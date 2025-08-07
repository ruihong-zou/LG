package com.example.demo;

import org.springframework.stereotype.Service;
import java.util.List;
import java.util.ArrayList;

/**
 * 批量翻译服务，支持API调用与本地模拟。
 */
@Service
public class TranslateService {

    // 参数常量集中管理
    private static final int BASE_WAIT_TIME = 180_000;    // 基础等待时间 3分钟
    private static final int PER_TEXT_WAIT_TIME = 10_000; // 每片段10秒
    private static final int PER_CHAR_WAIT_TIME = 200;    // 每字符200ms
    private static final int MIN_WAIT_TIME = 180_000;     // 最短3分钟
    private static final int MAX_WAIT_TIME = 600_000;     // 最长10分钟

    /**
     * 批量翻译文本列表
     */
    public List<String> batchTranslate(List<String> texts) {
        long batchStartTime = System.currentTimeMillis();
        java.text.SimpleDateFormat sdf = new java.text.SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS");

        System.out.println("=== 批量翻译开始 ===");
        System.out.println("开始时间: " + sdf.format(new java.util.Date(batchStartTime)));
        if (texts == null || texts.isEmpty()) {
            System.out.println("文本列表为空，直接返回");
            return new ArrayList<>();
        }
        
        System.out.println("输入文本数量: " + texts.size());
        
        // 分析输入文本的基本信息
        int totalChars = 0;
        int maxLength = 0;
        int minLength = Integer.MAX_VALUE;
        int emptyCount = 0;
        
        for (int i = 0; i < texts.size(); i++) {
            String text = texts.get(i);
            int length = (text != null) ? text.length() : 0;
            totalChars += length;
            maxLength = Math.max(maxLength, length);
            if (length > 0) {
                minLength = Math.min(minLength, length);
            }
            if (text == null || text.trim().isEmpty()) {
                emptyCount++;
            }
            
            // 打印前几个文本的详细信息
            // if (i < 3) {
            //     System.out.println(String.format("文本%d: 长度=%d, 内容=\"%s\"", 
            //         i + 1, length, text != null ? (text.length() > 50 ? text.substring(0, 50) + "..." : text) : "null"));
            // }
        }
        
        if (minLength == Integer.MAX_VALUE) minLength = 0;
        
        System.out.println("=== 输入文本统计信息 ===");
        System.out.println("总字符数: " + totalChars);
        System.out.println("平均长度: " + (texts.size() > 0 ? totalChars / texts.size() : 0));
        System.out.println("最长文本: " + maxLength + " 字符");
        System.out.println("最短文本: " + minLength + " 字符");
        System.out.println("空文本数量: " + emptyCount);

        // 构建JSON输入，逐一清理特殊字符
        System.out.println("=== 构建JSON输入 ===");
        com.alibaba.fastjson.JSONObject inputJson = new com.alibaba.fastjson.JSONObject();
        com.alibaba.fastjson.JSONArray textsArray = new com.alibaba.fastjson.JSONArray();
        
        int cleanedCount = 0;
        for (int i = 0; i < texts.size(); i++) {
            String originalText = texts.get(i);
            String cleanedText = cleanTextForJson(originalText);
            textsArray.add(cleanedText);
            
            if (!java.util.Objects.equals(originalText, cleanedText)) {
                cleanedCount++;
                if (cleanedCount <= 3) { // 只打印前3个清理示例
                    System.out.println(String.format("文本%d清理: \"%s\" -> \"%s\"", 
                        i + 1, 
                        originalText != null ? (originalText.length() > 30 ? originalText.substring(0, 30) + "..." : originalText) : "null",
                        cleanedText.length() > 30 ? cleanedText.substring(0, 30) + "..." : cleanedText));
                }
            }
        }
        
        if (cleanedCount > 0) {
            System.out.println("总共清理了 " + cleanedCount + " 个文本中的特殊字符");
        } else {
            System.out.println("所有文本无需清理");
        }
        
        inputJson.put("texts", textsArray);

        String jsonInput;
        try {
            System.out.println("=== JSON序列化 ===");
            jsonInput = inputJson.toJSONString();
            System.out.println("JSON序列化成功");
            System.out.println("JSON长度: " + jsonInput.length() + " 字符");
            System.out.println("JSON预览: " + (jsonInput.length() > 200 ? jsonInput.substring(0, 200) + "..." : jsonInput));
        } catch (Exception e) {
            System.err.println("=== JSON序列化失败，执行应急清理 ===");
            System.err.println("序列化失败原因: " + e.getMessage());
            
            // 应急清理
            textsArray.clear();
            for (String text : texts) {
                String emergencyCleanedText = text == null ? "" : text.replaceAll("[^\\u4e00-\\u9fa5a-zA-Z0-9\\s\\.,;:!?()\\[\\]{}\\-_=+*/@#$%&]", "");
                textsArray.add(emergencyCleanedText);
            }
            inputJson.put("texts", textsArray);
            jsonInput = inputJson.toJSONString();
            System.out.println("应急清理后JSON序列化成功");
            System.out.println("应急清理后JSON长度: " + jsonInput.length() + " 字符");
        }

        // 计算预计等待时间
        int expectedWaitTime = calculateWaitTime(texts.size(), totalChars);

        // 调用Kimi API，失败则自动走模拟翻译
        try {
            System.out.println("=== 开始调用Kimi API ===");
            System.out.println("输入JSON: " + jsonInput);
            System.out.println("输入片段数量: " + texts.size());
            System.out.println("预计等待时间: " + expectedWaitTime + "ms");
            
            String translatedResponse = Kimi.robustTranslateToEnglish(jsonInput);
            
            // 详细分析API返回内容
            System.out.println("=== API返回内容分析 ===");
            System.out.println("原始返回内容: " + translatedResponse);
            System.out.println("返回内容长度: " + (translatedResponse != null ? translatedResponse.length() : 0));
            System.out.println("返回内容类型检查: " + (translatedResponse != null && translatedResponse.trim().startsWith("{") ? "JSON格式" : "非JSON格式"));
            
            com.alibaba.fastjson.JSONObject responseJson = com.alibaba.fastjson.JSONObject.parseObject(translatedResponse);
            System.out.println("JSON解析成功，包含字段: " + responseJson.keySet());
            
            if (!responseJson.containsKey("translations")) {
                System.out.println("错误: API返回缺少translations字段");
                System.out.println("实际包含的字段: " + responseJson.keySet());
                throw new Exception("API返回缺少translations字段");
            }
            
            com.alibaba.fastjson.JSONArray translationsArray = responseJson.getJSONArray("translations");
            System.out.println("translations数组长度: " + translationsArray.size());
            System.out.println("期望的数组长度: " + texts.size());
            
            // 打印每个翻译结果的详细信息
            // System.out.println("=== 翻译结果详细分析 ===");
            // for (int i = 0; i < translationsArray.size(); i++) {
            //     String translation = translationsArray.getString(i);
            //     System.out.println(String.format("第%d个翻译 - 原文长度: %d, 译文长度: %d", 
            //         i + 1, 
            //         (i < texts.size() ? texts.get(i).length() : 0), 
            //         (translation != null ? translation.length() : 0)));
            //     System.out.println(String.format("  原文: %s", i < texts.size() ? texts.get(i) : "超出范围"));
            //     System.out.println(String.format("  译文: %s", translation));
            // }
            
            if (translationsArray.size() != texts.size()) {
                System.out.println("错误: 翻译数量不匹配");
                System.out.println("输入文本数量: " + texts.size());
                System.out.println("返回翻译数量: " + translationsArray.size());
                throw new Exception("翻译数量不匹配");
            }
            
            List<String> result = new ArrayList<>();
            for (int i = 0; i < translationsArray.size(); i++) {
                result.add(translationsArray.getString(i));
            }
            long batchEndTime = System.currentTimeMillis();
            long totalTime = batchEndTime - batchStartTime;
            // 统计翻译结果的总字符数
            int resultTotalChars = result.stream().mapToInt(s -> s != null ? s.length() : 0).sum();
            // 打印总结
            System.out.println("=== 翻译结果总结 ===");
            System.out.println("字段数（翻译条数）: " + result.size());
            System.out.println("总字符数: " + resultTotalChars);
            System.out.println("总用时: " + totalTime + " ms");
            System.out.printf("=== 批量翻译完成（总耗时: %d ms）===\n", totalTime);
            System.out.println("最终返回结果数量: " + result.size());
            return result;
        } catch (Exception e) {
            System.err.println("=== 批量翻译失败详细分析 ===");
            System.err.println("失败原因: " + e.getMessage());
            System.err.println("异常类型: " + e.getClass().getSimpleName());
            if (e.getCause() != null) {
                System.err.println("根本原因: " + e.getCause().getMessage());
            }
            
            // 打印调用栈的前几行以便调试
            System.err.println("调用栈信息:");
            StackTraceElement[] stackTrace = e.getStackTrace();
            for (int i = 0; i < Math.min(5, stackTrace.length); i++) {
                System.err.println("  " + stackTrace[i].toString());
            }
            
            System.err.println("=== 切换到本地模拟翻译 ===");
            System.err.println("输入文本数量: " + texts.size());
            return simulateBatchTranslate(texts);
        }
    }

    /**
     * 清理JSON中可能导致解析失败的特殊字符
     */
    private String cleanTextForJson(String text) {
        if (text == null) return "";
        return text.replace("\\", "\\\\")
                   .replace("\"", "\\\"")
                   .replace("\r", "\\r")
                   .replace("\n", "\\n")
                   .replace("\t", "\\t")
                   .replace("\b", "\\b")
                   .replace("\f", "\\f")
                   .replace("->", "→")
                   .replace("<-", "←")
                   .replace("=>", "⇒")
                   .replace("<=", "⇐")
                   .replace("<", "&lt;")
                   .replace(">", "&gt;");
    }

    /**
     * 根据片段数量和总字符数计算等待时间（ms）
     */
    private int calculateWaitTime(int textCount, int totalChars) {
        int waitTime = BASE_WAIT_TIME + (textCount * PER_TEXT_WAIT_TIME) + (totalChars * PER_CHAR_WAIT_TIME);
        return Math.max(MIN_WAIT_TIME, Math.min(MAX_WAIT_TIME, waitTime));
    }

    /**
     * 本地模拟翻译，适用于API异常时
     */
    private List<String> simulateBatchTranslate(List<String> texts) {
        List<String> result = new ArrayList<>();
        for (String text : texts) {
            result.add(simulateTranslate(text));
        }
        return result;
    }

    /**
     * 模拟单条翻译，中文包裹[模拟翻译]，英文包裹[Simulated]
     */
    private String simulateTranslate(String text) {
        String core = text == null ? "" : text.replace("\r", "").replace("\n", "");
        return containsChinese(core) ? "[模拟翻译]" + core + "[模拟翻译]" : "[Simulated]" + core + "[Simulated]";
    }

    /**
     * 判断文本是否包含中文
     */
    private boolean containsChinese(String text) {
        return text != null && text.matches(".*[\\u4e00-\\u9fa5].*");
    }
}
