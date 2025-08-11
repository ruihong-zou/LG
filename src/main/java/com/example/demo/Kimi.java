package com.example.demo;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import okhttp3.*;

import java.io.IOException;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.*;

@Slf4j
public final class Kimi {

    // ===== 基础配置（建议用 .env 注入） =====
    private static final String API_KEY = System.getenv().getOrDefault("MOONSHOT_API_KEY", "");
    private static final String CHAT_COMPLETION_URL = System.getenv().getOrDefault(
            "MOONSHOT_API_URL", "https://api.moonshot.cn/v1/chat/completions");

    // 账号配额（Tier3 可在 .env 设置；TPD=0 表示无限）
    private static final int RPM_LIMIT = getEnvInt("MOONSHOT_RPM", 5000);
    private static final int TPM_LIMIT = getEnvInt("MOONSHOT_TPM", 384000);
    private static final int TPD_LIMIT = getEnvInt("MOONSHOT_TPD", 0);

    // 显式设置的 max_tokens（网关按 prompt + max_tokens 计入 TPM/TPD；也影响是否 length 截断）
    private static final int MAX_COMPLETION_TOKENS = getEnvInt("MOONSHOT_MAX_COMPLETION", 1500);

    // 并发上限（不“串行”卡死；按机器核数/网速调；不填默认 32）
    private static final int CONCURRENCY_LIMIT = getEnvInt("MOONSHOT_CONCURRENCY", 32);

    private static int getEnvInt(String k, int d){
        try { return Integer.parseInt(System.getenv().getOrDefault(k, String.valueOf(d))); }
        catch(Exception e){ return d; }
    }

    private Kimi() {}

    // 全局限流器（并发=CONCURRENCY_LIMIT，守护 RPM/TPM/TPD）
    private static final RateLimiter LIMITER = new RateLimiter(RPM_LIMIT, TPM_LIMIT, TPD_LIMIT, CONCURRENCY_LIMIT);

    public enum Direction { ZH2EN, EN2ZH }

    @NoArgsConstructor @AllArgsConstructor @Data
    public static class MoonshotMessage { private String role; private String content; }

    /** 智能翻译（严格 texts→translations 对齐；length 直接抛异常给上层切批） */
    public static String robustTranslate(String jsonTexts, Direction direction) throws IOException {
        if (API_KEY.isEmpty()) throw new IOException("MOONSHOT_API_KEY 未配置");
        if (jsonTexts == null || jsonTexts.trim().isEmpty()) return "{\"translations\":[]}";

        int expectedCount = 0;
        try {
            com.alibaba.fastjson.JSONObject inputJson = com.alibaba.fastjson.JSONObject.parseObject(jsonTexts);
            if (inputJson.containsKey("texts")) expectedCount = inputJson.getJSONArray("texts").size();
        } catch (Exception ignore) {}

        final String systemPrompt = direction == Direction.ZH2EN
                ? "你是一名专业翻译助手。请逐一独立翻译 JSON 数组 texts 中每个片段为英文："
                  + "1) 保持顺序与数量严格与输入一致；空字符串输出空字符串；"
                  + "2) 任何包含中文汉字的片段（哪怕只有一个字/单个标签/序号/量词），必须译为自然流畅的英文表达，不得回显中文；"
                  + "3) 仅对纯标点/表情/特殊符号可原样返回；"
                  + "4) 输出必须是紧凑单行 JSON：{\"translations\":[...]}，无换行/无解释，并正确转义。"
                : "你是一名专业翻译助手。请逐一独立翻译 JSON 数组 texts 中每个片段为中文："
                  + "1) 保持顺序与数量严格与输入一致；空字符串输出空字符串；"
                  + "2) 任何包含英文字符的片段（包括单词/单个字母/标签/序号），必须译为自然流畅的中文表达，不得回显英文；"
                  + "3) 仅对纯标点/表情/特殊符号可原样返回；"
                  + "4) 输出必须是紧凑单行 JSON：{\"translations\":[...]}，无换行/无解释，并正确转义。";

        final String userPrompt = expectedCount > 0
                ? String.format("请翻译 JSON 中 %d 个片段，输出 translations 数组与输入 texts 数量一致：%s", expectedCount, jsonTexts)
                : "请翻译以下 JSON 格式内容，输出 translations 与 texts 数量一致：" + jsonTexts;

        List<MoonshotMessage> messages = new ArrayList<>(2);
        messages.add(new MoonshotMessage("system", systemPrompt));
        messages.add(new MoonshotMessage("user", userPrompt));

        for (int attempt = 0; attempt < 3; attempt++) {
            String resp = chatNoStreamWithFinishReason("moonshot-v1-32k", messages);
            String finish = parseFinishReason(resp);
            if ("length".equals(finish)) throw new IOException("API输出被截断（finish_reason=length）");

            String contentJson = parseContentString(resp);
            if (isValidAndMatchedCount(contentJson, expectedCount)) return contentJson;

            int actual = -1;
            try {
                com.alibaba.fastjson.JSONObject result = com.alibaba.fastjson.JSONObject.parseObject(contentJson);
                if (result.containsKey("translations")) {
                    com.alibaba.fastjson.JSONArray arr = result.getJSONArray("translations");
                    if (arr != null) actual = arr.size();
                }
            } catch (Exception ignore) {}
            String fix = String.format(
                    "你刚才输出的 translations 数量为 %d，与输入 texts 数量（%d）不一致。请严格一一对应、顺序不变，输出紧凑单行 JSON，数量必须为 %d，不得包含多余内容或换行。",
                    actual, expectedCount, expectedCount);
            messages.add(new MoonshotMessage("user", fix));
        }
        throw new IOException("多次自动修正后仍不一致，请拆分更小批次或检查输入。");
    }

    /** 底层对话（按 prompt+max_tokens 预占限流，显式 max_tokens） */
    public static String chatNoStreamWithFinishReason(String model, List<MoonshotMessage> messages) throws IOException {
        int promptTokensEst = estimatePromptTokensForMessages(messages);
        int budget = promptTokensEst + MAX_COMPLETION_TOKENS;
        LIMITER.beforeRequest(budget);
        try {
            cn.hutool.json.JSONObject payload = new cn.hutool.json.JSONObject()
                    .putOpt("model", model)
                    .putOpt("messages", messages)
                    .putOpt("stream", false)
                    .putOpt("response_format", new cn.hutool.json.JSONObject().putOpt("type", "json_object"))
                    .putOpt("max_tokens", MAX_COMPLETION_TOKENS);
            String requestBody = payload.toString();

            OkHttpClient client = new OkHttpClient.Builder()
                    .connectTimeout(60, java.util.concurrent.TimeUnit.SECONDS)
                    .writeTimeout(120, java.util.concurrent.TimeUnit.SECONDS)
                    .readTimeout(240, java.util.concurrent.TimeUnit.SECONDS)
                    .callTimeout(300, java.util.concurrent.TimeUnit.SECONDS)
                    .build();

            Request req = new Request.Builder()
                    .url(CHAT_COMPLETION_URL)
                    .post(RequestBody.create(requestBody, MediaType.get("application/json")))
                    .addHeader("Authorization", "Bearer " + API_KEY)
                    .build();

            try (Response resp = client.newCall(req).execute()) {
                if (resp.body() == null) throw new IOException("空响应体");
                String body = resp.body().string();
                if (resp.code() >= 400) throw new IOException("HTTP " + resp.code() + ": " + body);
                return body;
            }
        } finally {
            LIMITER.afterRequest();
        }
    }

    // ===== 解析与校验 =====
    private static String parseFinishReason(String responseJsonStr) {
        try {
            com.alibaba.fastjson.JSONObject responseJson = com.alibaba.fastjson.JSONObject.parseObject(responseJsonStr);
            com.alibaba.fastjson.JSONArray choices = responseJson.getJSONArray("choices");
            if (choices != null && !choices.isEmpty()) return choices.getJSONObject(0).getString("finish_reason");
        } catch (Exception ignore) {}
        return null;
    }

    private static String parseContentString(String responseJsonStr) {
        try {
            com.alibaba.fastjson.JSONObject responseJson = com.alibaba.fastjson.JSONObject.parseObject(responseJsonStr);
            com.alibaba.fastjson.JSONArray choices = responseJson.getJSONArray("choices");
            if (choices != null && !choices.isEmpty()) {
                String content = choices.getJSONObject(0).getJSONObject("message").getString("content");
                return stripJsonFence(content);
            }
        } catch (Exception ignore) {}
        return "";
    }

    private static String stripJsonFence(String s) {
        if (s == null) return "";
        String t = s.trim();
        if (t.startsWith("```") ) {
            int firstNl = t.indexOf('\n');
            if (firstNl >= 0) {
                String body = t.substring(firstNl + 1);
                int lastFence = body.lastIndexOf("```");
                if (lastFence >= 0) body = body.substring(0, lastFence);
                return body.trim();
            }
        }
        return t;
    }

    private static boolean isValidAndMatchedCount(String contentJson, int expectedCount) {
        if (contentJson == null || !contentJson.trim().startsWith("{")) return false;
        try {
            com.alibaba.fastjson.JSONObject result = com.alibaba.fastjson.JSONObject.parseObject(contentJson);
            if (!result.containsKey("translations")) return false;
            com.alibaba.fastjson.JSONArray arr = result.getJSONArray("translations");
            return arr != null && arr.size() == expectedCount;
        } catch (Exception e) { return false; }
    }

    // ===== 估算器（仅估算 prompt，用于限流预算） =====
    private static int estimatePromptTokensForMessages(List<MoonshotMessage> messages) {
        int input = 0;
        for (MoonshotMessage m : messages) input += estimateTokensForText(m.getContent());
        return input + 200; // 留出消息结构/role等开销
    }
    private static int estimateTokensForText(String s) {
        if (s == null || s.isEmpty()) return 0;
        int cjk = 0, other = 0;
        for (int i = 0; i < s.length(); i++) {
            char ch = s.charAt(i);
            if ((ch >= '\u4E00' && ch <= '\u9FFF') || (ch >= '\u3400' && ch <= '\u4DBF')) cjk++; else other++;
        }
        return cjk + (int)Math.ceil(other / 4.0);
    }

    // ===== 限流器（并发、RPM、TPM、TPD；TPD<=0 表示无限） =====
    private static final class RateLimiter {
        private final int rpmLimit; // 每分钟请求数
        private final int tpmLimit; // 每分钟 token
        private final int tpdLimit; // 每日 token（<=0 表示无限）

        private final java.util.concurrent.Semaphore concurrency;
        private final Deque<Long> recentRequests = new ArrayDeque<>();

        private long minuteWindowStart = System.currentTimeMillis();
        private int tokensThisMinute = 0;

        private LocalDate day = LocalDate.now(ZoneId.systemDefault());
        private int tokensToday = 0;

        RateLimiter(int rpm, int tpm, int tpd, int concurrencyLimit) {
            this.rpmLimit = rpm; this.tpmLimit = tpm; this.tpdLimit = tpd;
            this.concurrency = new java.util.concurrent.Semaphore(Math.max(1, concurrencyLimit));
        }

        void beforeRequest(int requestedTokens) {
            acquireConcurrency();
            synchronized (this) {
                while (true) {
                    long now = System.currentTimeMillis();
                    while (!recentRequests.isEmpty() && now - recentRequests.peekFirst() >= 60_000) recentRequests.pollFirst();
                    if (now - minuteWindowStart >= 60_000) { minuteWindowStart = now; tokensThisMinute = 0; }

                    LocalDate today = LocalDate.now(ZoneId.systemDefault());
                    if (!today.equals(day)) { day = today; tokensToday = 0; }

                    boolean rpmOk = recentRequests.size() < rpmLimit;
                    boolean tpmOk = (tokensThisMinute + requestedTokens) <= tpmLimit;
                    boolean tpdOk = (tpdLimit <= 0) || ((tokensToday + requestedTokens) <= tpdLimit);

                    if (rpmOk && tpmOk && tpdOk) {
                        recentRequests.addLast(now);
                        tokensThisMinute += requestedTokens;
                        tokensToday += requestedTokens;
                        return;
                    }

                    long sleepMs = 250L;
                    if (!rpmOk) {
                        long oldest = recentRequests.peekFirst();
                        sleepMs = Math.max(sleepMs, 60_000 - (now - oldest) + 5);
                    }
                    if (!tpmOk) {
                        sleepMs = Math.max(sleepMs, 60_000 - (now - minuteWindowStart) + 5);
                    }
                    if (!tpdOk) {
                        releaseConcurrency();
                        throw new RuntimeException("超出当日可用tokens预算(TPD)，请次日再试或降低用量");
                    }
                    try { this.wait(sleepMs); } catch (InterruptedException ie) { Thread.currentThread().interrupt(); }
                }
            }
        }

        void afterRequest() {
            synchronized (this) {
                long now = System.currentTimeMillis();
                while (!recentRequests.isEmpty() && now - recentRequests.peekFirst() >= 60_000) recentRequests.pollFirst();
                if (now - minuteWindowStart >= 60_000) { minuteWindowStart = now; tokensThisMinute = 0; }
                LocalDate today = LocalDate.now(ZoneId.systemDefault());
                if (!today.equals(day)) { day = today; tokensToday = 0; }
                this.notifyAll();
            }
            releaseConcurrency();
        }

        private void acquireConcurrency() {
            try { concurrency.acquire(); }
            catch (InterruptedException e) { Thread.currentThread().interrupt(); throw new RuntimeException("并发信号量获取被中断", e); }
        }
        private void releaseConcurrency() { concurrency.release(); }
    }
}
