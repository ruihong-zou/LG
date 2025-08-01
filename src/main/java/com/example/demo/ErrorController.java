package com.example.demo;

import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import jakarta.servlet.http.HttpServletRequest;
import java.util.HashMap;
import java.util.Map;

@Controller
public class ErrorController implements org.springframework.boot.web.servlet.error.ErrorController {

    @RequestMapping("/error")
    @ResponseBody
    public ResponseEntity<Map<String, Object>> handleError(HttpServletRequest request) {
        Map<String, Object> errorDetails = new HashMap<>();
        
        Integer statusCode = (Integer) request.getAttribute("jakarta.servlet.error.status_code");
        String errorMessage = (String) request.getAttribute("jakarta.servlet.error.message");
        Exception exception = (Exception) request.getAttribute("jakarta.servlet.error.exception");
        
        errorDetails.put("status", statusCode != null ? statusCode : 500);
        errorDetails.put("error", "Internal Server Error");
        errorDetails.put("message", errorMessage != null ? errorMessage : "An unexpected error occurred");
        
        if (exception != null) {
            errorDetails.put("exception", exception.getClass().getSimpleName());
            errorDetails.put("details", exception.getMessage());
        }
        
        System.err.println("错误详情: " + errorDetails);
        if (exception != null) {
            exception.printStackTrace();
        }
        
        return ResponseEntity.status(statusCode != null ? statusCode : 500).body(errorDetails);
    }
}




// <!-- //11 -->



