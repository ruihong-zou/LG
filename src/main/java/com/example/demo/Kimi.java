package com.example.demo;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.json.JSONObject;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import okhttp3.*;
import java.io.IOException;
import java.util.*;

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

    public static String chatNoStream(String model, List<MoonshotMessage> messages) throws IOException {
        String requestBody = new JSONObject()
                .putOpt("model", model)
                .putOpt("messages", messages)
                .putOpt("stream", false)
                .toString();

        okhttp3.Request okhttpRequest = new Request.Builder()
                .url(CHAT_COMPLETION_URL)
                .post(RequestBody.create(requestBody, MediaType.get("application/json")))
                .addHeader("Authorization", "Bearer " + API_KEY)
                .build();

        Call call = new OkHttpClient().newCall(okhttpRequest);
        try (Response okhttpResponse = call.execute()) {
            String responseBody = okhttpResponse.body().string();
            if (responseBody == null || responseBody.isEmpty()) {
                return "";
            }

            com.alibaba.fastjson.JSONObject responseJson = com.alibaba.fastjson.JSONObject.parseObject(responseBody);
            String content = responseJson.getJSONArray("choices")
                    .getJSONObject(0)
                    .getJSONObject("message")
                    .getString("content");

            return content != null ? content : "";
        }
    }


    public static void main(String[] args) throws IOException {
        List<MoonshotMessage> messages = CollUtil.newArrayList(
                new MoonshotMessage("user", "你好，你是谁啊")
        );
        String res = chatNoStream("moonshot-v1-8k", messages);
        System.out.println(res);
    }
}
