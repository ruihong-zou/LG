package com.example.demo;

import com.aspose.words.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.ss.usermodel.CellType;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFTextShape;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.sl.extractor.SlideShowExtractor;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.sl.usermodel.SlideShowFactory;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hssf.extractor.ExcelExtractor;

import java.io.ByteArrayOutputStream;
import java.util.HashMap;

@RestController
@RequestMapping("/api")
public class DocumentController {
    
    @Autowired
    private DocumentProcessor documentProcessor;
    
    @GetMapping("/")
    public String home() {
        return "Office Document Processor is running! 📄✨";
    }
    
    // Aspose 处理方法
    @PostMapping("/aspose/process")
    public ResponseEntity<byte[]> processWithAspose(@RequestParam("file") MultipartFile file) throws Exception {
        String filename = file.getOriginalFilename();
        
        if (filename.endsWith(".xlsx") || filename.endsWith(".xls")) {
            // 使用POI处理Excel文件
            XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
            
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            workbook.write(out);
            workbook.close();
            
            return ResponseEntity.ok()
                    .header("Content-Disposition", "attachment; filename=aspose-processed.xlsx")
                    .body(out.toByteArray());
        } else {
            // 处理Word文件
            Document doc = new Document(file.getInputStream());
            
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            doc.save(out, SaveFormat.DOCX);
            
            return ResponseEntity.ok()
                    .header("Content-Disposition", "attachment; filename=aspose-processed.docx")
                    .body(out.toByteArray());
        }
    }
    
    // Apache POI 处理方法
    @PostMapping("/poi/process")
    public ResponseEntity<byte[]> processWithPOI(@RequestParam("file") MultipartFile file) throws Exception {
        try {
            System.out.println("开始处理文件: " + file.getOriginalFilename());
            String filename = file.getOriginalFilename().toLowerCase();
            
            if (filename.endsWith(".xlsx")) {
                return processExcelXLSX(file);
            } else if (filename.endsWith(".xls")) {
                return processExcelXLS(file);
            } else if (filename.endsWith(".pptx")) {
                return processPowerPointPPTX(file);
            } else if (filename.endsWith(".ppt")) {
                return processPowerPointPPT(file);
            } else if (filename.endsWith(".docx")) {
                return processWordDOCX(file);
            } else if (filename.endsWith(".doc")) {
                return processWordDOC(file);
            } else {
                throw new IllegalArgumentException("不支持的文件格式: " + filename);
            }
        } catch (Exception e) {
            System.err.println("处理文件时出错: " + e.getMessage());
            e.printStackTrace();
            throw e;
        }
    }
    
    private ResponseEntity<byte[]> processExcelXLSX(MultipartFile file) throws Exception {
        System.out.println("处理Excel XLSX文件 - 使用批量翻译");
        XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
        
        // 使用新的批量处理逻辑
        workbook = documentProcessor.processExcelDocument(workbook);
        
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        workbook.close();
        
        return ResponseEntity.ok()
                .header("Content-Disposition", "attachment; filename=batch-translated.xlsx")
                .body(out.toByteArray());
    }
    
    private ResponseEntity<byte[]> processExcelXLS(MultipartFile file) throws Exception {
        System.out.println("处理Excel XLS文件 - 使用批量翻译");
        
        try {
            // 尝试作为传统XLS格式处理
            HSSFWorkbook workbook = new HSSFWorkbook(file.getInputStream());
            workbook = documentProcessor.processExcelXLS(workbook);
            
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            workbook.write(out);
            workbook.close();
            
            return ResponseEntity.ok()
                    .header("Content-Disposition", "attachment; filename=batch-translated.xls")
                    .body(out.toByteArray());
                    
        } catch (org.apache.poi.poifs.filesystem.OfficeXmlFileException e) {
            // 如果是XML格式，说明实际是XLSX文件，使用XLSX处理逻辑
            System.out.println("检测到文件实际为XLSX格式，切换到XLSX处理逻辑");
            return processExcelXLSX(file);
        }
    }
    
    private ResponseEntity<byte[]> processWordDOCX(MultipartFile file) throws Exception {
        System.out.println("处理Word DOCX文件 - 使用批量翻译");
        XWPFDocument doc = new XWPFDocument(file.getInputStream());
        
        // 使用新的批量处理逻辑
        doc = documentProcessor.processWordDocument(doc);
        
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        doc.write(out);
        doc.close();
        
        return ResponseEntity.ok()
                .header("Content-Disposition", "attachment; filename=batch-translated.docx")
                .body(out.toByteArray());
    }
    
    private ResponseEntity<byte[]> processWordDOC(MultipartFile file) throws Exception {
        System.out.println("处理Word DOC文件 - 使用批量翻译");
        
        try {
            // 尝试作为传统DOC格式处理
            HWPFDocument doc = new HWPFDocument(file.getInputStream());
            
            doc = documentProcessor.processWordDOC(doc);
            
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            doc.write(out);
            doc.close();
            
            return ResponseEntity.ok()
                    .header("Content-Disposition", "attachment; filename=batch-translated.doc")
                    .body(out.toByteArray());
                
        } catch (IllegalArgumentException e) {
            // 如果是OOXML格式，说明实际是DOCX文件，使用DOCX处理逻辑
            if (e.getMessage().contains("OOXML")) {
                System.out.println("检测到文件实际为DOCX格式，切换到DOCX处理逻辑");
                return processWordDOCX(file);
            } else {
                throw e;
            }
        }
    }
    
    private ResponseEntity<byte[]> processPowerPointPPTX(MultipartFile file) throws Exception {
        System.out.println("处理PowerPoint PPTX文件 - 使用批量翻译");
        XMLSlideShow ppt = new XMLSlideShow(file.getInputStream());
        
        // 使用新的批量处理逻辑
        ppt = documentProcessor.processPowerPointPPTX(ppt);
        
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ppt.write(out);
        ppt.close();
        
        return ResponseEntity.ok()
                .header("Content-Disposition", "attachment; filename=batch-translated.pptx")
                .body(out.toByteArray());
    }
    
    private ResponseEntity<byte[]> processPowerPointPPT(MultipartFile file) throws Exception {
        System.out.println("处理PowerPoint PPT文件 - 使用批量翻译");
        
        try {
            HSLFSlideShow ppt = new HSLFSlideShow(file.getInputStream());
            
            // 使用新的批量处理逻辑
            ppt = documentProcessor.processPowerPointPPT(ppt);
            
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ppt.write(out);
            ppt.close();
            
            return ResponseEntity.ok()
                    .header("Content-Disposition", "attachment; filename=batch-translated.ppt")
                    .body(out.toByteArray());
                
        } catch (org.apache.poi.poifs.filesystem.OfficeXmlFileException e) {
            // 如果是XML格式，说明实际是PPTX文件，使用PPTX处理逻辑
            System.out.println("检测到文件实际为PPTX格式，切换到PPTX处理逻辑");
            return processPowerPointPPTX(file);
        }
    }
    
    // docx4j 处理方法
    @PostMapping("/docx4j/process")
    public ResponseEntity<byte[]> processWithDocx4j(@RequestParam("file") MultipartFile file) throws Exception {
        WordprocessingMLPackage wordPackage = WordprocessingMLPackage.load(file.getInputStream());
        
        // 获取主文档部分
        MainDocumentPart mainPart = wordPackage.getMainDocumentPart();
        

        
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        wordPackage.save(out);
        
        return ResponseEntity.ok()
                .header("Content-Disposition", "attachment; filename=docx4j-processed.docx")
                .body(out.toByteArray());
    }
}
