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

import java.io.ByteArrayOutputStream;
import java.util.HashMap;

@RestController
@RequestMapping("/api")
public class DocumentController {
    
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
        System.out.println("处理Excel XLSX文件");
        XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
        
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            XSSFSheet sheet = workbook.getSheetAt(i);
            for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                XSSFRow row = sheet.getRow(rowNum);
                if (row != null) {
                    for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
                        XSSFCell cell = row.getCell(cellNum);
                        if (cell != null && cell.getCellType() == CellType.STRING) {
                            String cellValue = cell.getStringCellValue();
                            if (cellValue != null && !cellValue.trim().isEmpty()) {
                                cell.setCellValue("[翻译]" + cellValue);
                            }
                        }
                    }
                }
            }
        }
        
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        workbook.close();
        
        return ResponseEntity.ok()
                .header("Content-Disposition", "attachment; filename=poi-processed.xlsx")
                .body(out.toByteArray());
    }
    
    private ResponseEntity<byte[]> processExcelXLS(MultipartFile file) throws Exception {
        System.out.println("处理Excel XLS文件");
        HSSFWorkbook workbook = new HSSFWorkbook(file.getInputStream());
        
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            HSSFSheet sheet = workbook.getSheetAt(i);
            for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                HSSFRow row = sheet.getRow(rowNum);
                if (row != null) {
                    for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
                        HSSFCell cell = row.getCell(cellNum);
                        if (cell != null && cell.getCellType() == CellType.STRING) {
                            String cellValue = cell.getStringCellValue();
                            if (cellValue != null && !cellValue.trim().isEmpty()) {
                                cell.setCellValue("[翻译]" + cellValue);
                            }
                        }
                    }
                }
            }
        }
        
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        workbook.close();
        
        return ResponseEntity.ok()
                .header("Content-Disposition", "attachment; filename=poi-processed.xls")
                .body(out.toByteArray());
    }
    
    private ResponseEntity<byte[]> processWordDOCX(MultipartFile file) throws Exception {
        System.out.println("处理Word DOCX文件");
        XWPFDocument doc = new XWPFDocument(file.getInputStream());
        
        for (XWPFParagraph paragraph : doc.getParagraphs()) {
            for (XWPFRun run : paragraph.getRuns()) {
                String text = run.getText(0);
                if (text != null && !text.trim().isEmpty()) {
                    run.setText("[翻译]" + text, 0);
                }
            }
        }
        
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        doc.write(out);
        doc.close();
        
        return ResponseEntity.ok()
                .header("Content-Disposition", "attachment; filename=poi-processed.docx")
                .body(out.toByteArray());
    }
    
    private ResponseEntity<byte[]> processWordDOC(MultipartFile file) throws Exception {
        System.out.println("处理Word DOC文件");
        HWPFDocument doc = new HWPFDocument(file.getInputStream());
        Range range = doc.getRange();
        
        // 获取文档文本并添加翻译标记
        String text = range.text();
        if (text != null && !text.trim().isEmpty()) {
            // 简单的文本替换处理
            String[] paragraphs = text.split("\r");
            StringBuilder processedText = new StringBuilder();
            for (String paragraph : paragraphs) {
                if (!paragraph.trim().isEmpty()) {
                    processedText.append("[翻译]").append(paragraph).append("\r");
                } else {
                    processedText.append(paragraph).append("\r");
                }
            }
            range.replaceText(text, processedText.toString());
        }
        
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        doc.write(out);
        doc.close();
        
        return ResponseEntity.ok()
                .header("Content-Disposition", "attachment; filename=poi-processed.doc")
                .body(out.toByteArray());
    }
    
    private ResponseEntity<byte[]> processPowerPointPPTX(MultipartFile file) throws Exception {
        System.out.println("处理PowerPoint PPTX文件");
        XMLSlideShow ppt = new XMLSlideShow(file.getInputStream());
        
        for (XSLFSlide slide : ppt.getSlides()) {
            for (XSLFShape shape : slide.getShapes()) {
                if (shape instanceof XSLFTextShape) {
                    XSLFTextShape textShape = (XSLFTextShape) shape;
                    String text = textShape.getText();
                    if (text != null && !text.trim().isEmpty()) {
                        textShape.setText("[翻译]" + text);
                    }
                }
            }
        }
        
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ppt.write(out);
        ppt.close();
        
        return ResponseEntity.ok()
                .header("Content-Disposition", "attachment; filename=poi-processed.pptx")
                .body(out.toByteArray());
    }
    
    private ResponseEntity<byte[]> processPowerPointPPT(MultipartFile file) throws Exception {
        System.out.println("处理PowerPoint PPT文件");
        HSLFSlideShow ppt = new HSLFSlideShow(file.getInputStream());
        
        for (HSLFSlide slide : ppt.getSlides()) {
            for (HSLFShape shape : slide.getShapes()) {
                if (shape instanceof HSLFTextShape) {
                    HSLFTextShape textShape = (HSLFTextShape) shape;
                    String text = textShape.getText();
                    if (text != null && !text.trim().isEmpty()) {
                        textShape.setText("[翻译]" + text);
                    }
                }
            }
        }
        
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ppt.write(out);
        ppt.close();
        
        return ResponseEntity.ok()
                .header("Content-Disposition", "attachment; filename=poi-processed.ppt")
                .body(out.toByteArray());
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










