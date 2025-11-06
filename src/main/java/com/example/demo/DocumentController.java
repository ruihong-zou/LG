package com.example.demo;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

@RestController
@RequestMapping("/api")
public class DocumentController {

    @Autowired
    private DocumentProcessor documentProcessor;
    @Autowired
    private OfficeConvertService officeConvertService;
    
    @GetMapping("/")
    public String home() {
        return "Office Document Processor is running! 📄✨";
    }
    
    // Apache POI 处理方法
@PostMapping("/process")
public ResponseEntity<byte[]> processWithPOI(
        @RequestParam("file") MultipartFile file,
        @RequestParam(value = "sourceLang", required = false, defaultValue = "auto") String sourceLang,
        @RequestParam(value = "targetLang", required = false, defaultValue = "en") String targetLang,
        @RequestParam(value = "userPrompt", required = false) String userPrompt
) throws Exception {
    try {
        System.out.println("开始处理文件: " + file.getOriginalFilename());
        String filename = file.getOriginalFilename().toLowerCase();

        if (filename.endsWith(".xlsx")) {
            return processExcelXLSX(file, targetLang, userPrompt);
        } else if (filename.endsWith(".xls")) {
            return processExcelXLS(file, targetLang, userPrompt);
        } else if (filename.endsWith(".pptx")) {
            return processPowerPointPPTX(file, targetLang, userPrompt);
        } else if (filename.endsWith(".ppt")) {
            return processPowerPointPPT(file, targetLang, userPrompt);
        } else if (filename.endsWith(".docx")) {
            return processWordDOCX(file, targetLang, userPrompt);
        } else if (filename.endsWith(".doc")) {
            return processWordDOC(file, targetLang, userPrompt);
        } else {
            throw new IllegalArgumentException("不支持的文件格式: " + filename);
        }

    } catch (Exception e) {
        System.err.println("处理文件时出错: " + e.getMessage());
        e.printStackTrace();
        throw e;
    }
}
    
    private ResponseEntity<byte[]> processExcelXLSX(MultipartFile file, String targetLang, String userPrompt) throws Exception {
        System.out.println("处理Excel XLSX文件 - 使用批量翻译");
        XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
        
        // 使用新的批量处理逻辑
        workbook = documentProcessor.processExcelDocument(workbook, targetLang, userPrompt);
        
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        workbook.close();
        
        return ResponseEntity.ok()
                .header("Content-Disposition", "attachment; filename=batch-translated.xlsx")
                .body(out.toByteArray());
    }
    
    private ResponseEntity<byte[]> processExcelXLS(MultipartFile file, String targetLang, String userPrompt) throws Exception {
        System.out.println("处理Excel XLS文件 - 使用批量翻译");
        
        try {
            // 尝试作为传统XLS格式处理
            HSSFWorkbook workbook = new HSSFWorkbook(file.getInputStream());
            workbook = documentProcessor.processExcelXLS(workbook, targetLang, userPrompt);
            
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            workbook.write(out);
            workbook.close();
            
            return ResponseEntity.ok()
                    .header("Content-Disposition", "attachment; filename=batch-translated.xls")
                    .body(out.toByteArray());
                    
        } catch (org.apache.poi.poifs.filesystem.OfficeXmlFileException e) {
            // 如果是XML格式，说明实际是XLSX文件，使用XLSX处理逻辑
            System.out.println("检测到文件实际为XLSX格式，切换到XLSX处理逻辑");
            return processExcelXLSX(file, targetLang, userPrompt);
        }
    }
    
    private ResponseEntity<byte[]> processWordDOCX(MultipartFile file, String targetLang, String userPrompt) throws Exception {
        System.out.println("处理Word DOCX文件 - 使用批量翻译");
        XWPFDocument doc = new XWPFDocument(file.getInputStream());
        
        // 使用新的批量处理逻辑
        doc = documentProcessor.processWordDocument(doc, targetLang, userPrompt);
        
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        doc.write(out);
        doc.close();
        
        return ResponseEntity.ok()
                .header("Content-Disposition", "attachment; filename=batch-translated.docx")
                .body(out.toByteArray());
    }
    
    private ResponseEntity<byte[]> processWordDOC(MultipartFile file, String targetLang, String userPrompt) throws Exception {
        System.out.println("处理Word DOC文件 - 先转DOCX翻译，最终仍输出DOC");

        // 1) 读入原始 .doc
        byte[] originalDoc = file.getBytes();

        // 2) .doc -> .docx （只用于中间处理）
        byte[] asDocx = officeConvertService.docToDocx(originalDoc);

        // 3) 在 .docx 上执行已有的翻译逻辑
        byte[] translatedDocx;
        try (XWPFDocument xdoc = new XWPFDocument(new ByteArrayInputStream(asDocx));
            ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            XWPFDocument translated = documentProcessor.processWordDocument(xdoc, targetLang, userPrompt);
            translated.write(out);
            translated.close();
            translatedDocx = out.toByteArray();
        }

        // 4) 将翻译后的 .docx -> .doc，保证输出扩展仍为 .doc
        byte[] finalDoc = officeConvertService.docxToDoc(translatedDocx);

        return ResponseEntity.ok()
                .header("Content-Disposition", "attachment; filename=batch-translated.doc")
                .body(finalDoc);
    }

    private ResponseEntity<byte[]> processPowerPointPPTX(MultipartFile file, String targetLang, String userPrompt) throws Exception {
        System.out.println("处理PowerPoint PPTX文件 - 使用批量翻译");
        XMLSlideShow ppt = new XMLSlideShow(file.getInputStream());
        
        // 使用新的批量处理逻辑
        ppt = documentProcessor.processPowerPointPPTX(ppt, targetLang, userPrompt);
        
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ppt.write(out);
        ppt.close();
        
        return ResponseEntity.ok()
                .header("Content-Disposition", "attachment; filename=batch-translated.pptx")
                .body(out.toByteArray());
    }
    
    private ResponseEntity<byte[]> processPowerPointPPT(MultipartFile file, String targetLang, String userPrompt) throws Exception {
        System.out.println("处理PowerPoint PPT文件 - 使用批量翻译");
        
        try {
            HSLFSlideShow ppt = new HSLFSlideShow(file.getInputStream());
            
            // 使用新的批量处理逻辑
            ppt = documentProcessor.processPowerPointPPT(ppt, targetLang, userPrompt);
            
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ppt.write(out);
            ppt.close();
            
            return ResponseEntity.ok()
                    .header("Content-Disposition", "attachment; filename=batch-translated.ppt")
                    .body(out.toByteArray());
                
        } catch (org.apache.poi.poifs.filesystem.OfficeXmlFileException e) {
            // 如果是XML格式，说明实际是PPTX文件，使用PPTX处理逻辑
            System.out.println("检测到文件实际为PPTX格式，切换到PPTX处理逻辑");
            return processPowerPointPPTX(file, targetLang, userPrompt);
        }
    }
    

}
