package com.example.demo;

import org.springframework.stereotype.Service;
import java.util.List;
import java.util.ArrayList;

@Service
public class TranslateService {
    
    public List<String> batchTranslate(List<String> texts) {
        System.out.println("批量翻译 " + texts.size() + " 个文本片段");
        List<String> translatedTexts = new ArrayList<>();
        
        for (String text : texts) {
            if (text != null && !text.trim().isEmpty()) {
                // 模拟翻译，并在控制台模拟没一段话的翻译
                String translated = simulateTranslate(text);
                translatedTexts.add(translated);
                System.out.println("原文: " + text + " -> 识别文字: " + translated);
            } else {
                translatedTexts.add(text);
            }
        }
        
        return translatedTexts;
    }
    
    private String simulateTranslate(String text) {

        // 先把所有回车和换行都去掉
        String core = text.replace("\r", "").replace("\n", "");
        // 不要 trim()——保持开头末尾空格（如果业务需要可再决定）

        // 简单的中文探测逻辑，模拟翻译场景
        if (containsChinese(core)) {
            return "[中文][中文]" + core + "[中文][中文]";
        } else {
            return "[无中文][无中文]" + core + "[无中文][无中文]";
        }
    }
    
    private boolean containsChinese(String text) {
        return text.matches(".*[\\u4e00-\\u9fa5].*");
    }
}