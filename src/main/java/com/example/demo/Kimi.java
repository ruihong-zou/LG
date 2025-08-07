package com.example.demo;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import okhttp3.*;
import java.io.IOException;
import java.util.*;

/**
 * Kimi大模型API调用工具类
 */
public class Kimi {

    private static final String API_KEY = "sk-C9xOUe9C8Ed4R37gwt45XGkrNuDsUZAbqpYd7vBkJGIaqu5L";
    private static final String CHAT_COMPLETION_URL = "https://api.moonshot.cn/v1/chat/completions";

    @NoArgsConstructor
    @AllArgsConstructor
    @Data
    @Builder
    public static class MoonshotMessage {
        private String role;
        private String content;
    }

    /**
     * 发送请求，返回moonshot完整响应JSON字符串（带finish_reason）
     */
    public static String chatNoStreamWithFinishReason(String model, List<MoonshotMessage> messages) throws IOException {
        System.out.println("=== Kimi API 对话开始 ===");
        
        // 打印对话内容详细分析
        System.out.println("=== 对话内容分析 ===");
        System.out.println("使用模型: " + model);
        System.out.println("消息总数: " + messages.size());
        
        for (int i = 0; i < messages.size(); i++) {
            MoonshotMessage msg = messages.get(i);
            System.out.println(String.format("消息%d [%s]: ", i + 1, msg.getRole()));
            String content = msg.getContent();
            if (content != null) {
                System.out.println("  长度: " + content.length() + " 字符");
                // 如果内容太长，只显示前200字符
                if (content.length() > 200) {
                    System.out.println("  内容: " + content.substring(0, 200) + "...[截断]");
                } else {
                    System.out.println("  内容: " + content);
                }
                
                // 分析消息类型和特征
                if ("system".equals(msg.getRole())) {
                    System.out.println("  类型: 系统提示词");
                } else if ("user".equals(msg.getRole())) {
                    System.out.println("  类型: 用户输入");
                    if (content.contains("{") && content.contains("texts")) {
                        System.out.println("  特征: 包含JSON翻译请求");
                    }
                    if (content.contains("翻译JSON中") && content.contains("个片段")) {
                        System.out.println("  特征: 批量翻译请求");
                    }
                    if (content.contains("输出translations数组数量与输入texts不一致")) {
                        System.out.println("  特征: 纠错追问消息");
                    }
                } else if ("assistant".equals(msg.getRole())) {
                    System.out.println("  类型: AI回复");
                }
            } else {
                System.out.println("  内容: null");
            }
            System.out.println();
        }
        
        String requestBody = new cn.hutool.json.JSONObject()
                .putOpt("model", model)
                .putOpt("messages", messages)
                .putOpt("stream", false)
                .putOpt("response_format", new cn.hutool.json.JSONObject().putOpt("type", "json_object"))
                .toString();
        
        System.out.println("=== 请求详情 ===");
        System.out.println("请求体长度: " + requestBody.length() + " 字符");
        System.out.println("请求体内容: " + (requestBody.length() > 500 ? requestBody.substring(0, 500) + "...[截断]" : requestBody));

        OkHttpClient client = new OkHttpClient.Builder()
                .connectTimeout(300, java.util.concurrent.TimeUnit.SECONDS)
                .writeTimeout(300, java.util.concurrent.TimeUnit.SECONDS)
                .readTimeout(900, java.util.concurrent.TimeUnit.SECONDS)
                .callTimeout(1200, java.util.concurrent.TimeUnit.SECONDS)
                .build();

        Request okhttpRequest = new Request.Builder()
                .url(CHAT_COMPLETION_URL)
                .post(RequestBody.create(requestBody, MediaType.get("application/json")))
                .addHeader("Authorization", "Bearer " + API_KEY)
                .build();

        try (Response okhttpResponse = client.newCall(okhttpRequest).execute()) {
            String responseBody = okhttpResponse.body().string();
            
            System.out.println("=== API响应分析 ===");
            System.out.println("响应状态码: " + okhttpResponse.code());
            System.out.println("响应头信息:");
            okhttpResponse.headers().forEach(header -> 
                System.out.println("  " + header.getFirst() + ": " + header.getSecond()));
            
            if (responseBody == null || responseBody.isEmpty()) {
                System.out.println("错误: API返回空响应");
                throw new IOException("API返回空响应");
            }
            
            System.out.println("响应体长度: " + responseBody.length() + " 字符");
            System.out.println("完整响应内容: " + responseBody);
            
            // 解析并分析响应结构
            try {
                com.alibaba.fastjson.JSONObject responseJson = com.alibaba.fastjson.JSONObject.parseObject(responseBody);
                System.out.println("=== 响应结构分析 ===");
                System.out.println("响应JSON包含字段: " + responseJson.keySet());
                
                if (responseJson.containsKey("choices")) {
                    com.alibaba.fastjson.JSONArray choices = responseJson.getJSONArray("choices");
                    System.out.println("choices数组长度: " + choices.size());
                    
                    if (choices.size() > 0) {
                        com.alibaba.fastjson.JSONObject firstChoice = choices.getJSONObject(0);
                        System.out.println("第一个choice包含字段: " + firstChoice.keySet());
                        
                        if (firstChoice.containsKey("finish_reason")) {
                            String finishReason = firstChoice.getString("finish_reason");
                            System.out.println("finish_reason: " + finishReason);
                            if ("length".equals(finishReason)) {
                                System.out.println("警告: 响应被截断!");
                            }
                        }
                        
                        if (firstChoice.containsKey("message")) {
                            com.alibaba.fastjson.JSONObject message = firstChoice.getJSONObject("message");
                            System.out.println("message字段: " + message.keySet());
                            if (message.containsKey("content")) {
                                String content = message.getString("content");
                                System.out.println("AI回复内容长度: " + (content != null ? content.length() : 0));
                                System.out.println("AI回复内容: " + content);
                            }
                        }
                    }
                }
                
                if (responseJson.containsKey("usage")) {
                    com.alibaba.fastjson.JSONObject usage = responseJson.getJSONObject("usage");
                    System.out.println("Token使用统计: " + usage);
                }
            } catch (Exception e) {
                System.out.println("响应JSON解析失败: " + e.getMessage());
            }
            
            System.out.println("=== API对话完成 ===");
            return responseBody;
        }
    }

    /**
     * 智能校验和自动追问修正的中译英接口
     */
    public static String robustTranslateToEnglish(String text) throws IOException {
        return robustTranslate(text, true);
    }
    /**
     * 智能校验和自动追问修正的英译中接口
     */
    public static String robustTranslateToChinese(String text) throws IOException {
        return robustTranslate(text, false);
    }

    /**
     * 通用：自动检测translations数量并递归纠错，遇到length立即终止
     * @param text 输入JSON文本
     * @param zhToEn true=中译英, false=英译中
     */
    public static String robustTranslate(String text, boolean zhToEn) throws IOException {
        System.out.println("=== 智能翻译开始 ===");
        System.out.println("翻译方向: " + (zhToEn ? "中文→英文" : "英文→中文"));
        System.out.println("输入文本: " + text);
        
        if (text == null || text.trim().isEmpty()) {
            System.out.println("输入为空，直接返回");
            return text;
        }

        int expectedCount = 0;
        try {
            com.alibaba.fastjson.JSONObject inputJson = com.alibaba.fastjson.JSONObject.parseObject(text);
            if (inputJson.containsKey("texts")) {
                expectedCount = inputJson.getJSONArray("texts").size();
                System.out.println("解析到期望翻译数量: " + expectedCount);
            }
        } catch (Exception e) {
            System.out.println("输入JSON解析失败: " + e.getMessage());
        }

        String systemPrompt =
                zhToEn
                        ? "你是一名专业翻译助手。请严格逐一独立翻译用户提供的JSON数组texts中每个片段，禁止合并、忽略、重排，空字符串也要输出空翻译，输出数组translations和输入texts数量严格一致。输出必须为紧凑单行JSON，例如：{\"translations\":[\"result1\",\"result2\"]}，不得包含换行符、格式化符、注释或说明，特殊字符需正确转义。"
                        : "你是一名专业翻译助手。请严格逐一独立翻译用户提供的JSON数组texts中每个片段为中文，禁止合并、忽略、重排，空字符串也要输出空翻译，输出数组translations和输入texts数量严格一致。输出必须为紧凑单行JSON，例如：{\"translations\":[\"结果1\",\"结果2\"]}，不得包含换行符、格式化符、注释或说明，特殊字符需正确转义。";
        String userContent = expectedCount > 0
                ? String.format("请翻译JSON中%d个片段，输出translations数组与输入texts数量一致，%s", expectedCount, text)
                : "请翻译以下JSON格式内容，输出translations与texts数量一致：" + text;

        List<MoonshotMessage> messages = new ArrayList<>();
        messages.add(new MoonshotMessage("system", systemPrompt));
        messages.add(new MoonshotMessage("user", userContent));
        
        System.out.println("=== 初始对话设置完成 ===");
        System.out.println("系统提示词长度: " + systemPrompt.length());
        System.out.println("用户消息长度: " + userContent.length());

        for (int attempt = 0; attempt < 3; attempt++) {
            System.out.println("=== 第" + (attempt + 1) + "轮对话开始 ===");
            String responseJson = chatNoStreamWithFinishReason("moonshot-v1-8k", messages);

            String finishReason = parseFinishReason(responseJson);
            System.out.println("finish_reason: " + finishReason);
            if ("length".equals(finishReason)) {
                System.out.println("API输出被截断，终止重试");
                throw new IOException("API输出被截断（finish_reason=length），请分批翻译！");
            }

            String contentJson = parseContentString(responseJson);
            System.out.println("AI返回内容: " + contentJson);
            
            // 验证返回内容
            boolean isValid = isValidAndMatchedCount(contentJson, expectedCount);
            System.out.println("内容验证结果: " + (isValid ? "通过" : "失败"));
            
            if (isValid) {
                System.out.println("=== 智能翻译成功完成 ===");
                return contentJson;
            } else {
                // 分析失败原因
                analyzeValidationFailure(contentJson, expectedCount);
                
                // 统计AI刚才返回的实际数量
                int actualCount = -1;
                try {
                    com.alibaba.fastjson.JSONObject resultJson = com.alibaba.fastjson.JSONObject.parseObject(contentJson);
                    if (resultJson.containsKey("translations")) {
                        com.alibaba.fastjson.JSONArray arr = resultJson.getJSONArray("translations");
                        if (arr != null) actualCount = arr.size();
                    }
                } catch (Exception ignored) {}

                // 追问修正，只反馈数量
                String fixPrompt = String.format(
                        "你刚才输出的translations数组数量为%d，与输入texts数量（%d）不一致，请严格逐一对应，输出紧凑单行JSON，数量必须为%d，顺序不变，且不得有任何多余内容或换行。",
                        actualCount, expectedCount, expectedCount
                );
                System.out.println("=== 添加纠错消息 ===");
                System.out.println("纠错提示: " + fixPrompt);
                // messages.add(new MoonshotMessage("assistant", contentJson)); // 这一行可以去掉
                messages.add(new MoonshotMessage("user", fixPrompt));
            }
        }
        System.out.println("=== 智能翻译失败 ===");
        throw new IOException("多次自动修正后，AI输出依然不符，请检查输入分片或内容是否合适。");
    }

    /** 提取"finish_reason" */
    private static String parseFinishReason(String responseJsonStr) {
        System.out.println("=== 解析finish_reason ===");
        try {
            com.alibaba.fastjson.JSONObject responseJson = com.alibaba.fastjson.JSONObject.parseObject(responseJsonStr);
            com.alibaba.fastjson.JSONArray choices = responseJson.getJSONArray("choices");
            if (choices != null && choices.size() > 0) {
                String finishReason = choices.getJSONObject(0).getString("finish_reason");
                System.out.println("提取到finish_reason: " + finishReason);
                return finishReason;
            } else {
                System.out.println("choices数组为空或不存在");
            }
        } catch (Exception e) {
            System.out.println("解析finish_reason失败: " + e.getMessage());
        }
        return null;
    }

    /** 提取content字段 */
    private static String parseContentString(String responseJsonStr) {
        System.out.println("=== 解析content字段 ===");
        try {
            com.alibaba.fastjson.JSONObject responseJson = com.alibaba.fastjson.JSONObject.parseObject(responseJsonStr);
            com.alibaba.fastjson.JSONArray choices = responseJson.getJSONArray("choices");
            if (choices != null && choices.size() > 0) {
                String content = choices.getJSONObject(0).getJSONObject("message").getString("content");
                System.out.println("提取到content长度: " + (content != null ? content.length() : 0));
                System.out.println("提取到content内容: " + content);
                return content;
            } else {
                System.out.println("choices数组为空或不存在");
            }
        } catch (Exception e) {
            System.out.println("解析content失败: " + e.getMessage());
        }
        return "";
    }

    /** 校验content字符串是否为紧凑JSON、且translations数量与输入一致 */
    private static boolean isValidAndMatchedCount(String contentJson, int expectedCount) {
        System.out.println("=== 验证翻译结果 ===");
        System.out.println("期望数量: " + expectedCount);
        System.out.println("待验证内容: " + contentJson);
        
        if (contentJson == null || !contentJson.trim().startsWith("{")) {
            System.out.println("验证失败: 不是JSON格式");
            return false;
        }
        
        try {
            com.alibaba.fastjson.JSONObject resultJson = com.alibaba.fastjson.JSONObject.parseObject(contentJson);
            System.out.println("JSON解析成功，包含字段: " + resultJson.keySet());
            
            if (!resultJson.containsKey("translations")) {
                System.out.println("验证失败: 缺少translations字段");
                return false;
            }
            
            com.alibaba.fastjson.JSONArray arr = resultJson.getJSONArray("translations");
            if (arr == null) {
                System.out.println("验证失败: translations字段为null");
                return false;
            }
            
            System.out.println("实际翻译数量: " + arr.size());
            boolean isValid = arr.size() == expectedCount;
            System.out.println("数量匹配: " + isValid);
            
            // 打印翻译内容详情
            if (arr.size() > 0) {
                System.out.println("翻译结果预览:");
                for (int i = 0; i < Math.min(3, arr.size()); i++) {
                    System.out.println("  " + (i + 1) + ": " + arr.getString(i));
                }
                if (arr.size() > 3) {
                    System.out.println("  ... 还有" + (arr.size() - 3) + "个结果");
                }
            }
            
            return isValid;
        } catch (Exception e) {
            System.out.println("验证失败: JSON解析异常 - " + e.getMessage());
            return false;
        }
    }

    /**
     * 分析验证失败的具体原因
     */
    private static void analyzeValidationFailure(String contentJson, int expectedCount) {
        System.out.println("=== 验证失败原因分析 ===");
        
        if (contentJson == null) {
            System.out.println("失败类型: 内容为null");
            return;
        }
        
        if (!contentJson.trim().startsWith("{")) {
            System.out.println("失败类型: 非JSON格式");
            System.out.println("内容开头: " + contentJson.substring(0, Math.min(50, contentJson.length())));
            return;
        }
        
        try {
            com.alibaba.fastjson.JSONObject resultJson = com.alibaba.fastjson.JSONObject.parseObject(contentJson);
            
            if (!resultJson.containsKey("translations")) {
                System.out.println("失败类型: 缺少translations字段");
                System.out.println("实际包含字段: " + resultJson.keySet());
                return;
            }
            
            com.alibaba.fastjson.JSONArray arr = resultJson.getJSONArray("translations");
            if (arr == null) {
                System.out.println("失败类型: translations字段为null");
                return;
            }
            
            System.out.println("失败类型: 数量不匹配");
            System.out.println("期望数量: " + expectedCount);
            System.out.println("实际数量: " + arr.size());
            System.out.println("差异: " + (arr.size() - expectedCount));
            
            if (arr.size() > expectedCount) {
                System.out.println("可能原因: AI生成了额外的翻译结果");
            } else {
                System.out.println("可能原因: AI遗漏了部分翻译结果");
            }
            
        } catch (Exception e) {
            System.out.println("失败类型: JSON解析错误");
            System.out.println("解析错误: " + e.getMessage());
        }
    }

    // === 兼容旧接口 ===

    /**
     * 只尝试一次，不纠错的中译英
     */
    public static String translateToEnglish(String text) throws IOException {
        return translateRaw(text, true);
    }
    /**
     * 只尝试一次，不纠错的英译中
     */
    public static String translateToChinese(String text) throws IOException {
        return translateRaw(text, false);
    }

    /** 只尝试一次的通用接口，兼容原有写法 */
    public static String translateRaw(String text, boolean zhToEn) throws IOException {
        if (text == null || text.trim().isEmpty()) return text;

        int expectedCount = 0;
        try {
            com.alibaba.fastjson.JSONObject inputJson = com.alibaba.fastjson.JSONObject.parseObject(text);
            if (inputJson.containsKey("texts")) {
                expectedCount = inputJson.getJSONArray("texts").size();
            }
        } catch (Exception ignored) {}

        String systemPrompt =
                zhToEn
                        ? "你是一名专业翻译助手。请严格逐一独立翻译用户提供的JSON数组texts中每个片段，禁止合并、忽略、重排，空字符串也要输出空翻译，输出数组translations和输入texts数量严格一致。输出必须为紧凑单行JSON，例如：{\"translations\":[\"result1\",\"result2\"]}，不得包含换行符、格式化符、注释或说明，特殊字符需正确转义。"
                        : "你是一名专业翻译助手。请严格逐一独立翻译用户提供的JSON数组texts中每个片段为中文，禁止合并、忽略、重排，空字符串也要输出空翻译，输出数组translations和输入texts数量严格一致。输出必须为紧凑单行JSON，例如：{\"translations\":[\"结果1\",\"结果2\"]}，不得包含换行符、格式化符、注释或说明，特殊字符需正确转义。";
        String userContent = expectedCount > 0
                ? String.format("请翻译JSON中%d个片段，输出translations数组与输入texts数量一致，%s", expectedCount, text)
                : "请翻译以下JSON格式内容，输出translations与texts数量一致：" + text;

        List<MoonshotMessage> messages = new ArrayList<>();
        messages.add(new MoonshotMessage("system", systemPrompt));
        messages.add(new MoonshotMessage("user", userContent));

        String responseJson = chatNoStreamWithFinishReason("moonshot-v1-8k", messages);
        return parseContentString(responseJson);
    }
}