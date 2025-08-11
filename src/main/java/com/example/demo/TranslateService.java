package com.example.demo;

import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.concurrent.*;
import java.util.stream.Collectors;

@Slf4j
@Service
public class TranslateService {

    // ===== 参数（可用 .env 覆盖） =====
    // Tier3：把单次预算与批大小默认放宽；仍可用 ENV 覆盖
    private static final int MAX_TOKENS_PER_REQUEST = getEnvInt("TRANSLATE_MAX_TOKENS_PER_REQUEST", 30000); // prompt + max_tokens
    private static final int MAX_ITEMS_PER_BATCH     = getEnvInt("TRANSLATE_MAX_ITEMS_PER_BATCH", 500);
    private static final int EST_PROMPT_OVERHEAD     = 300; // system+user+JSON结构开销

    private static final int MAX_COMPLETION_TOKENS   = getEnvInt("MOONSHOT_MAX_COMPLETION", 1500);
    private static final int COMPLETION_SAFETY_MARGIN= getEnvInt("TRANSLATE_COMPLETION_MARGIN", 50);

    private static final int PARALLELISM             = getEnvInt("MOONSHOT_CONCURRENCY", 32); // 与 Kimi 限流器一致

    private static final Kimi.Direction DEFAULT_DIR =
            parseDir(System.getenv().getOrDefault("TRANSLATE_DIRECTION", "ZH2EN"));

    // —— 策略/诊断开关（沿用你已有的） ——
    private static final boolean TRIVIAL_PASSTHROUGH =
            Boolean.parseBoolean(System.getenv().getOrDefault("TRANSLATE_TRIVIAL_PASSTHROUGH","true"));
    private static final boolean LOG_SEGMENT_DECISIONS =
            Boolean.parseBoolean(System.getenv().getOrDefault("TRANSLATE_LOG_SEGMENTS","false"));
    private static final int LOG_PREVIEW_CHARS = getEnvInt("TRANSLATE_LOG_PREVIEW_CHARS", 30);

    private static int getEnvInt(String k, int d){ try { return Integer.parseInt(System.getenv().getOrDefault(k, String.valueOf(d))); } catch(Exception e){ return d; } }
    private static Kimi.Direction parseDir(String v){ return "EN2ZH".equalsIgnoreCase(v) ? Kimi.Direction.EN2ZH : Kimi.Direction.ZH2EN; }

    // —— 输出估算：修正为“按输入 token 的倍数”，不再用 prompt 比例（避免过度切分） ——
    private static double outputFactor(Kimi.Direction d){ return d == Kimi.Direction.ZH2EN ? 1.20 : 0.90; }

    /** 默认方向（ENV: TRANSLATE_DIRECTION），批量翻译 */
    public List<String> batchTranslate(List<String> texts) { return batchTranslate(texts, DEFAULT_DIR); }

    /** 指定方向的批量翻译（并发跑批 & 保序） */
    public List<String> batchTranslate(List<String> texts, Kimi.Direction direction) {
        if (texts == null || texts.isEmpty()) return new ArrayList<>();

        long t0 = System.currentTimeMillis();
        log.info("translate: size={}, dir={}, parallelism={}", texts.size(), direction, Math.max(1, PARALLELISM));

        // 清理 + 默认 CJK 空格归一化
        final int N = texts.size();
        List<String> cleaned = new ArrayList<>(N);
        for (String s : texts) cleaned.add(cleanForJson(s));

        // 规划批次（返回一组连续区间）
        List<Range> plan = planBatches(cleaned, direction);
        log.info("planned batches: {}, avg size≈{}", plan.size(), N / Math.max(1, plan.size()));

        // 并发执行所有批次，并把结果写回固定数组，保证全局顺序
        String[] out = new String[N];
        if (plan.size() == 1) {
            // 小优化：单批串行即可
            Range r = plan.get(0);
            List<String> partRes = translateOneBatchWithAutoSplit(cleaned.subList(r.start, r.end), direction);
            for (int i=0;i<partRes.size();i++) out[r.start+i] = partRes.get(i);
        } else {
            int poolSize = Math.min(PARALLELISM, plan.size());
            ExecutorService exec = Executors.newFixedThreadPool(poolSize);
            List<Callable<Void>> jobs = new ArrayList<>(plan.size());
            for (Range r : plan) {
                jobs.add(() -> {
                    List<String> partRes = translateOneBatchWithAutoSplit(cleaned.subList(r.start, r.end), direction);
                    for (int i = 0; i < partRes.size(); i++) out[r.start + i] = partRes.get(i);
                    return null;
                });
            }
            try {
                List<Future<Void>> fs = exec.invokeAll(jobs);
                for (Future<Void> f : fs) f.get(); // 触发异常传播
            } catch (InterruptedException ie) {
                Thread.currentThread().interrupt();
                log.warn("parallel interrupted, falling back to sequential...");
                // 兜底串行
                for (Range r : plan) {
                    List<String> partRes = translateOneBatchWithAutoSplit(cleaned.subList(r.start, r.end), direction);
                    for (int i=0;i<partRes.size();i++) out[r.start+i] = partRes.get(i);
                }
            } catch (ExecutionException ee) {
                log.warn("parallel execution error: {}", ee.getMessage());
                // 兜底串行
                for (Range r : plan) {
                    List<String> partRes = translateOneBatchWithAutoSplit(cleaned.subList(r.start, r.end), direction);
                    for (int i=0;i<partRes.size();i++) out[r.start+i] = partRes.get(i);
                }
            } finally {
                exec.shutdownNow();
            }
        }

        List<String> results = new ArrayList<>(N);
        for (String s : out) results.add(Objects.requireNonNullElse(s, ""));
        log.info("done in {} ms", (System.currentTimeMillis() - t0));
        return results;
    }

    // —— 单批执行（length / 数量不一致 → 自动细分重试） ——
    private List<String> translateOneBatchWithAutoSplit(List<String> part, Kimi.Direction direction) {
        final int n = part.size();
        String[] out = new String[n];
        List<Integer> idx = new ArrayList<>();
        List<String> pay = new ArrayList<>();

        int trivialCount = 0;
        for (int i = 0; i < n; i++) {
            String s = part.get(i);
            if (TRIVIAL_PASSTHROUGH && isTrivialSegment(s)) {
                out[i] = s == null ? "" : s;
                trivialCount++;
                if (LOG_SEGMENT_DECISIONS) log.debug("seg[{}] PASSTHRU: \"{}\"", i, preview(s));
            } else {
                idx.add(i); pay.add(s);
                if (LOG_SEGMENT_DECISIONS) log.debug("seg[{}] → MODEL : \"{}\"", i, preview(s));
            }
        }
        if (trivialCount > 0) log.info("trivial bypassed in this batch: {}", trivialCount);
        if (pay.isEmpty()) return java.util.Arrays.asList(out);

        com.alibaba.fastjson.JSONObject input = new com.alibaba.fastjson.JSONObject();
        input.put("texts", toFastJsonArray(pay));
        String json;
        try { json = input.toJSONString(); }
        catch (Exception e) {
            log.warn("json serialize failed, emergency clean: {}", e.getMessage());
            List<String> emergency = pay.stream().map(this::emergencyClean).collect(Collectors.toList());
            input.put("texts", toFastJsonArray(emergency));
            json = input.toJSONString();
        }

        try {
            String resp = Kimi.robustTranslate(json, direction);
            com.alibaba.fastjson.JSONObject obj = com.alibaba.fastjson.JSONObject.parseObject(resp);
            com.alibaba.fastjson.JSONArray arr = obj.getJSONArray("translations");
            if (arr == null || arr.size() != pay.size()) throw new RuntimeException("size mismatch");

            int sameCount = 0;
            for (int k = 0; k < pay.size(); k++) {
                String in = pay.get(k);
                String outStr = arr.getString(k);
                boolean same = safeEqualsTrim(in, outStr);
                if (same) sameCount++;
                out[idx.get(k)] = outStr;

                if (LOG_SEGMENT_DECISIONS) {
                    log.debug("seg[{}] MODEL-OUT: in=\"{}\" | out=\"{}\"{}",
                            idx.get(k), preview(in), preview(outStr),
                            same ? "  (UNCHANGED)" : "");
                }
            }
            if (LOG_SEGMENT_DECISIONS && sameCount > 0) {
                log.info("model returned unchanged items in this batch: {}", sameCount);
            }
            return java.util.Arrays.asList(out);

        } catch (Exception e) {
            String msg = String.valueOf(e.getMessage());
            boolean needSplit =
                msg != null && (
                    msg.contains("finish_reason=length") ||
                    msg.contains("仍不一致") ||
                    msg.contains("not equal") ||
                    msg.contains("size mismatch")
                );

            if (needSplit) {
                log.warn("need split & retry, reason={}", msg);
                if (part.size() <= 1) {
                    log.warn("single item still not fixable, fallback to simulate: \"{}\"", preview(part.get(0)));
                    return simulateBatch(part);
                }
                int mid = part.size() / 2;
                List<String> left  = translateOneBatchWithAutoSplit(part.subList(0, mid), direction);
                List<String> right = translateOneBatchWithAutoSplit(part.subList(mid, part.size()), direction);
                List<String> merged = new ArrayList<>(left.size() + right.size());
                merged.addAll(left); merged.addAll(right);
                return merged;
            }

            log.warn("batch failed (simulate this batch): {} | first=\"{}\"", e.toString(), preview(part.get(0)));
            return simulateBatch(part);
        }
    }

    // ===== 批次规划：用“输入token总和”估算输出，再对照 max_tokens 分批 =====
    private static final class Range { final int start, end; Range(int s,int e){ this.start=s; this.end=e; } }

    private List<Range> planBatches(List<String> cleaned, Kimi.Direction direction) {
        List<Range> plan = new ArrayList<>();
        int curStart = 0;
        int curCount = 0;

        int promptTokens = EST_PROMPT_OVERHEAD;     // 当前批的 prompt token 总估
        int outputTokens = 0;                       // 当前批的“预计输出”总估（= sum(input_tokens * factor)）
        double of = outputFactor(direction);

        for (int i = 0; i < cleaned.size(); i++) {
            String s = cleaned.get(i);
            int t = estimateTokens(s);
            int projectedPrompt = promptTokens + t;
            int projectedOutput = (int)Math.ceil(outputTokens + t * of) + 2; // +2 粗略结构开销

            boolean exceedCount = (curCount + 1) > MAX_ITEMS_PER_BATCH;
            boolean exceedPromptPlusMax = (projectedPrompt + MAX_COMPLETION_TOKENS) > MAX_TOKENS_PER_REQUEST;
            boolean exceedCompletion = projectedOutput > (MAX_COMPLETION_TOKENS - COMPLETION_SAFETY_MARGIN);

            if (curCount > 0 && (exceedCount || exceedPromptPlusMax || exceedCompletion)) {
                plan.add(new Range(curStart, i));
                // reset
                curStart = i; curCount = 0;
                promptTokens = EST_PROMPT_OVERHEAD;
                outputTokens = 0;
                // 重新计算当前项作为新批的第一个
                projectedPrompt = EST_PROMPT_OVERHEAD + t;
                projectedOutput = (int)Math.ceil(t * of) + 2;
            }

            // 把当前项放入批
            curCount++;
            promptTokens = projectedPrompt;
            outputTokens = (int)Math.ceil(outputTokens + t * of);

            // 单条超预算：强制单条一批
            if ((promptTokens + MAX_COMPLETION_TOKENS) > MAX_TOKENS_PER_REQUEST
                || (outputTokens + 2) > (MAX_COMPLETION_TOKENS - COMPLETION_SAFETY_MARGIN)) {
                plan.add(new Range(curStart, i + 1));
                curStart = i + 1; curCount = 0;
                promptTokens = EST_PROMPT_OVERHEAD;
                outputTokens = 0;
            }
        }
        if (curStart < cleaned.size()) plan.add(new Range(curStart, cleaned.size()));

        return plan;
    }

    // ===== 估算器（粗略） =====
    private int estimateTokens(String s){
        if (s == null || s.isEmpty()) return 0;
        int cjk=0, other=0; for (int i=0;i<s.length();i++){ char ch=s.charAt(i); if ((ch>='\u4E00'&&ch<='\u9FFF')||(ch>='\u3400'&&ch<='\u4DBF')) cjk++; else other++; }
        return cjk + (int)Math.ceil(other/4.0);
    }

    // ===== 清理辅助（含默认开启的 CJK 空格归一化） =====
    private String cleanForJson(String s){
        if (s==null) return "";
        String t = stripControls(removeUnpairedSurrogates(removeBOM(s)))
                .replace("\r\n","\n").replace("\r","\n");
        t = normalizeForTranslation(t);
        return t;
    }
    private String emergencyClean(String s){ return stripControls(removeUnpairedSurrogates(removeBOM(Objects.requireNonNullElse(s, "")))); }
    private String removeBOM(String s){ return (!s.isEmpty() && s.charAt(0)=='\uFEFF') ? s.substring(1) : s; }
    private String stripControls(String s){ return s.replaceAll("[\\p{Cntrl}&&[^\\r\\n\\t]]"," "); }
    private String removeUnpairedSurrogates(String s){ StringBuilder sb=new StringBuilder(s.length()); for(int i=0;i<s.length();i++){ char ch=s.charAt(i); if(Character.isHighSurrogate(ch)){ if(i+1<s.length()&&Character.isLowSurrogate(s.charAt(i+1))){ sb.append(ch).append(s.charAt(++i)); } } else if(!Character.isLowSurrogate(ch)){ sb.append(ch);} } return sb.toString(); }

    // —— 归一化：去掉“汉字 与 汉字”之间的异常空白（多轮迭代，直到收敛） ——
    private String normalizeForTranslation(String s){
        if (s == null || s.isEmpty()) return s;
        String t = s.replace('\u00A0',' ').replaceAll("[ \\t]{2,}", " ").trim();
        String prev;
        do {
            prev = t;
            t = t.replaceAll("([\\p{IsHan}])\\s+([\\p{IsHan}])", "$1$2");
        } while (!t.equals(prev));
        return t;
    }

    // ===== “微小/纯标点”判定（更保守的直通策略） =====
    private boolean isTrivialSegment(String s) {
        if (s == null) return true;
        String t = s.trim();
        if (t.isEmpty()) return true;

        // 一旦包含“汉字”（CJK统一表意），必须送模型
        if (hasCJKIdeograph(t)) return false;

        // 纯空白/纯标点（ASCII 或 CJK 标点）→ 直通
        boolean punctOnly = t.codePoints().allMatch(ch ->
            Character.isWhitespace(ch) || isAsciiPunct(ch) || isCjkPunct(ch)
        );
        if (punctOnly) return true;

        // 极短 ASCII：仅单字符字母/数字 才直通（如 "A"、"3"）
        if (t.length() == 1 && t.codePoints().allMatch(ch -> ch < 128 && Character.isLetterOrDigit(ch))) {
            return true;
        }

        return false;
    }

    private boolean isAsciiPunct(int ch) { return ch < 128 && String.valueOf((char) ch).matches("\\p{Punct}"); }
    private boolean isCjkPunct(int ch) { return "、，。！？：；…—·《》〈〉“”‘’（）【】".indexOf((char) ch) >= 0; }
    private boolean hasCJKIdeograph(String s) {
        return s.codePoints().anyMatch(ch -> {
            Character.UnicodeBlock b = Character.UnicodeBlock.of(ch);
            return b == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS
                || b == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS_EXTENSION_A
                || b == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS_EXTENSION_B
                || b == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS_EXTENSION_C
                || b == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS_EXTENSION_D
                || b == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS_EXTENSION_E
                || b == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS_EXTENSION_F
                || b == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS_EXTENSION_G
                || b == Character.UnicodeBlock.CJK_COMPATIBILITY_IDEOGRAPHS;
        });
    }

    // ===== fastjson JSONArray 构造辅助 =====
    private com.alibaba.fastjson.JSONArray toFastJsonArray(List<String> list){
        com.alibaba.fastjson.JSONArray arr=new com.alibaba.fastjson.JSONArray();
        if(list!=null){ for(String s:list) arr.add(s); }
        return arr;
    }

    // ===== 本地模拟（仅兜底） =====
    private List<String> simulateBatch(List<String> texts){ List<String> out=new ArrayList<>(texts.size()); for(String s:texts) out.add(simulateOne(s)); return out; }
    private String simulateOne(String s){ String core = s==null?"":s.replace("\r","").replace("\n",""); return containsChinese(core)?"[模拟翻译]"+core+"[模拟翻译]":"[Simulated]"+core+"[Simulated]"; }
    private boolean containsChinese(String s){ return s!=null && s.matches(".*[\\u4e00-\\u9fa5].*"); }

    // ===== 诊断辅助 =====
    private String preview(String s) {
        if (s == null) return "null";
        String t = s.replace("\n"," ").replace("\r"," ");
        return t.length() <= LOG_PREVIEW_CHARS ? t : t.substring(0, LOG_PREVIEW_CHARS) + "...";
    }
    private boolean safeEqualsTrim(String a, String b) {
        String x = a == null ? "" : a.trim();
        String y = b == null ? "" : b.trim();
        return x.equals(y);
    }
}
