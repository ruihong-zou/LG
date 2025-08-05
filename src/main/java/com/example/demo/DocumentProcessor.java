package com.example.demo;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.hslf.usermodel.*;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import java.io.ByteArrayOutputStream;
import java.io.ByteArrayInputStream;

@Component
public class DocumentProcessor {
    
    @Autowired
    private TranslateService translateService;
    
    // 文档结构信息类
    public static class TextElement {
        public String text;
        public String type;
        public Map<String, Object> position;
        
        public TextElement(String text, String type, Map<String, Object> position) {
            this.text = text;
            this.type = type;
            this.position = position;
        }
    }
    
    // 1. Word DOCX文档处理
    public XWPFDocument processWordDocument(XWPFDocument doc) throws Exception {
        System.out.println("开始批量处理Word文档");
        
        List<TextElement> elements = extractWordTexts(doc);
        System.out.println("提取到 " + elements.size() + " 个文本元素");
        
        List<String> texts = new ArrayList<>();
        for (TextElement element : elements) {
            texts.add(element.text);
        }
        List<String> translatedTexts = translateService.batchTranslate(texts);
        
        restoreWordTexts(doc, elements, translatedTexts);
        
        return doc;
    }
    
    private List<TextElement> extractWordTexts(XWPFDocument doc) {
        List<TextElement> elements = new ArrayList<>();
        
        for (int i = 0; i < doc.getParagraphs().size(); i++) {
            XWPFParagraph paragraph = doc.getParagraphs().get(i);
            for (int j = 0; j < paragraph.getRuns().size(); j++) {
                XWPFRun run = paragraph.getRuns().get(j);
                String text = run.getText(0);
                if (text != null && !text.trim().isEmpty()) {
                    Map<String, Object> position = new HashMap<>();
                    position.put("paragraphIndex", i);
                    position.put("runIndex", j);
                    elements.add(new TextElement(text, "run", position));
                }
            }
        }
        return elements;
    }
    
    private void restoreWordTexts(XWPFDocument doc, List<TextElement> elements, List<String> translatedTexts) {
        for (int i = 0; i < elements.size(); i++) {
            TextElement element = elements.get(i);
            int paragraphIndex = (Integer) element.position.get("paragraphIndex");
            int runIndex = (Integer) element.position.get("runIndex");
            
            XWPFRun run = doc.getParagraphs().get(paragraphIndex).getRuns().get(runIndex);
            run.setText(translatedTexts.get(i), 0);
        }
    }
    
    // 2. Excel XLSX文档处理
    public XSSFWorkbook processExcelDocument(XSSFWorkbook workbook) throws Exception {
        System.out.println("开始批量处理Excel文档");
        
        List<TextElement> elements = extractExcelTexts(workbook);
        System.out.println("提取到 " + elements.size() + " 个单元格文本");
        
        List<String> texts = new ArrayList<>();
        for (TextElement element : elements) {
            texts.add(element.text);
        }
        List<String> translatedTexts = translateService.batchTranslate(texts);
        
        restoreExcelTexts(workbook, elements, translatedTexts);
        return workbook;
    }
    
    private List<TextElement> extractExcelTexts(XSSFWorkbook workbook) {
        List<TextElement> elements = new ArrayList<>();
        
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
            for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                XSSFRow row = sheet.getRow(rowNum);
                if (row != null) {
                    for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
                        XSSFCell cell = row.getCell(cellNum);
                        if (cell != null && cell.getCellType() == CellType.STRING) {
                            String text = cell.getStringCellValue();
                            if (text != null && !text.trim().isEmpty()) {
                                Map<String, Object> position = new HashMap<>();
                                position.put("sheetIndex", sheetIndex);
                                position.put("rowIndex", rowNum);
                                position.put("cellIndex", cellNum);
                                elements.add(new TextElement(text, "cell", position));
                            }
                        }
                    }
                }
            }
        }
        return elements;
    }
    
    private void restoreExcelTexts(XSSFWorkbook workbook, List<TextElement> elements, List<String> translatedTexts) {
        for (int i = 0; i < elements.size(); i++) {
            TextElement element = elements.get(i);
            int sheetIndex = (Integer) element.position.get("sheetIndex");
            int rowIndex = (Integer) element.position.get("rowIndex");
            int cellIndex = (Integer) element.position.get("cellIndex");
            
            XSSFCell cell = workbook.getSheetAt(sheetIndex).getRow(rowIndex).getCell(cellIndex);
            cell.setCellValue(translatedTexts.get(i));
        }
    }
    
    // 3. PowerPoint PPTX处理
    public XMLSlideShow processPowerPointPPTX(XMLSlideShow ppt) throws Exception {
        System.out.println("开始批量处理PowerPoint PPTX文档");
        
        List<TextElement> elements = extractPPTXTexts(ppt);
        System.out.println("提取到 " + elements.size() + " 个文本元素");
        
        List<String> texts = new ArrayList<>();
        for (TextElement element : elements) {
            texts.add(element.text);
        }
        List<String> translatedTexts = translateService.batchTranslate(texts);
        
        restorePPTXTexts(ppt, elements, translatedTexts);
        return ppt;
    }
    
    private List<TextElement> extractPPTXTexts(XMLSlideShow ppt) {
        List<TextElement> elements = new ArrayList<>();
        
        for (int slideIndex = 0; slideIndex < ppt.getSlides().size(); slideIndex++) {
            XSLFSlide slide = ppt.getSlides().get(slideIndex);
            for (int shapeIndex = 0; shapeIndex < slide.getShapes().size(); shapeIndex++) {
                XSLFShape shape = slide.getShapes().get(shapeIndex);
                if (shape instanceof XSLFTextShape) {
                    XSLFTextShape textShape = (XSLFTextShape) shape;
                    
                    for (int paragraphIndex = 0; paragraphIndex < textShape.getTextParagraphs().size(); paragraphIndex++) {
                        XSLFTextParagraph paragraph = textShape.getTextParagraphs().get(paragraphIndex);
                        for (int runIndex = 0; runIndex < paragraph.getTextRuns().size(); runIndex++) {
                            XSLFTextRun run = paragraph.getTextRuns().get(runIndex);
                            String text = run.getRawText();
                            if (text != null && !text.trim().isEmpty()) {
                                Map<String, Object> position = new HashMap<>();
                                position.put("slideIndex", slideIndex);
                                position.put("shapeIndex", shapeIndex);
                                position.put("paragraphIndex", paragraphIndex);
                                position.put("runIndex", runIndex);
                                elements.add(new TextElement(text, "textRun", position));
                            }
                        }
                    }
                }
            }
        }
        return elements;
    }
    
    private void restorePPTXTexts(XMLSlideShow ppt, List<TextElement> elements, List<String> translatedTexts) {
        for (int i = 0; i < elements.size(); i++) {
            TextElement element = elements.get(i);
            int slideIndex = (Integer) element.position.get("slideIndex");
            int shapeIndex = (Integer) element.position.get("shapeIndex");
            int paragraphIndex = (Integer) element.position.get("paragraphIndex");
            int runIndex = (Integer) element.position.get("runIndex");
            
            XSLFSlide slide = ppt.getSlides().get(slideIndex);
            XSLFShape shape = slide.getShapes().get(shapeIndex);
            if (shape instanceof XSLFTextShape) {
                XSLFTextShape textShape = (XSLFTextShape) shape;
                XSLFTextParagraph paragraph = textShape.getTextParagraphs().get(paragraphIndex);
                XSLFTextRun run = paragraph.getTextRuns().get(runIndex);
                
                run.setText(translatedTexts.get(i));
            }
        }
    }
    
    // 4. Word DOC处理
    public HWPFDocument processWordDOC(HWPFDocument doc) throws Exception {
        System.out.println("开始批量处理Word DOC文档");
        
            List<TextElement> elements = extractDOCTexts(doc);
            System.out.println("提取到 " + elements.size() + " 个文本元素");
            
            List<String> texts = new ArrayList<>();
            for (TextElement element : elements) {
                texts.add(element.text);
            }
            List<String> translatedTexts = translateService.batchTranslate(texts);
            
            restoreDOCTexts(doc, elements, translatedTexts);
            return doc;
    }
    
    private List<TextElement> extractDOCTexts(HWPFDocument doc) {
        List<TextElement> elements = new ArrayList<>();
        
            Range range = doc.getRange();
            for (int i = 0; i < range.numParagraphs(); i++) {
                org.apache.poi.hwpf.usermodel.Paragraph paragraph = range.getParagraph(i);
                String text = paragraph.text();
                if (text != null && !text.trim().isEmpty()) {
                    Map<String, Object> position = new HashMap<>();
                    position.put("paragraphIndex", i);
                    position.put("originalText", text);
                    elements.add(new TextElement(text, "paragraph", position));
                }
            }

        return elements;
    }
    
    private void restoreDOCTexts(HWPFDocument doc, List<TextElement> elements, List<String> translatedTexts) {
        System.out.println("=== restoreDOCTexts 开始执行 ===");
        
        try {
            System.out.println("步骤1: 获取Range对象...");
            Range range = doc.getRange();
            System.out.println("步骤1: Range对象获取成功，范围长度: " + range.text().length());
            
            System.out.println("步骤2: 按照Apache POI官方方式进行文本替换...");
            
            // 按照官方测试用例的方式，直接在range上进行替换
            for (int i = 0; i < elements.size(); i++) {
                TextElement element = elements.get(i);
                String originalText = (String) element.position.get("originalText");
                String translatedText = translatedTexts.get(i);
                
                System.out.println("步骤3." + i + ": 替换文本 [" + originalText.substring(0, Math.min(30, originalText.length())) + "...]");
                
                try {
                    // 使用官方推荐的方式：直接在range上替换
                    range.replaceText(originalText, translatedText);
                    System.out.println("步骤3." + i + ": 替换成功");
                } catch (Exception e) {
                    System.err.println("步骤3." + i + ": 替换失败 - " + e.getMessage());
                    // 继续处理下一个，不中断整个流程
                }
            }
            
            System.out.println("DOC文档文本替换完成");
            
        } catch (Exception e) {
            System.err.println("=== DOC文本替换出错 ===");
            System.err.println("错误信息: " + e.getMessage());
            e.printStackTrace();
        }
        
        System.out.println("=== restoreDOCTexts 执行结束 ===");
    }
    
    // 5. Excel XLS处理
    public HSSFWorkbook processExcelXLS(HSSFWorkbook workbook) throws Exception {
        System.out.println("开始批量处理Excel XLS文档");
        
        ExcelExtractor extractor = new ExcelExtractor(workbook);
        extractor.setFormulasNotResults(false);
        extractor.setIncludeSheetNames(false);
        String text = extractor.getText();
        extractor.close();
        
        if (text != null && !text.trim().isEmpty()) {
            String[] lines = text.split("\n");
            List<String> texts = new ArrayList<>();
            for (String line : lines) {
                if (line != null && !line.trim().isEmpty()) {
                    texts.add(line.trim());
                }
            }
            
            if (!texts.isEmpty()) {
                List<String> translatedTexts = translateService.batchTranslate(texts);
                restoreXLSTextsFromExtractor(workbook, texts, translatedTexts);
            }
        }
        
        return workbook;
    }
    
    private void restoreXLSTextsFromExtractor(HSSFWorkbook workbook, List<String> originalTexts, List<String> translatedTexts) {
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            HSSFSheet sheet = workbook.getSheetAt(sheetIndex);
            for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                HSSFRow row = sheet.getRow(rowNum);
                if (row != null) {
                    for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
                        HSSFCell cell = row.getCell(cellNum);
                        if (cell != null && cell.getCellType() == CellType.STRING) {
                            String cellValue = cell.getStringCellValue().trim();
                            if (cellValue != null && !cellValue.isEmpty()) {
                                for (int i = 0; i < originalTexts.size(); i++) {
                                    if (originalTexts.get(i).equals(cellValue)) {
                                        cell.setCellValue(translatedTexts.get(i));
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    
    // 6. PowerPoint PPT处理
    public HSLFSlideShow processPowerPointPPT(HSLFSlideShow ppt) throws Exception {
        System.out.println("开始批量处理PowerPoint PPT文档");
        
        List<TextElement> elements = extractPPTTexts(ppt);
        System.out.println("提取到 " + elements.size() + " 个文本元素");
        
        List<String> texts = new ArrayList<>();
        for (TextElement element : elements) {
            texts.add(element.text);
        }
        List<String> translatedTexts = translateService.batchTranslate(texts);
        
        restorePPTTexts(ppt, elements, translatedTexts);
        return ppt;
    }
    
    private List<TextElement> extractPPTTexts(HSLFSlideShow ppt) {
        List<TextElement> elements = new ArrayList<>();
        
        for (int slideIndex = 0; slideIndex < ppt.getSlides().size(); slideIndex++) {
            HSLFSlide slide = ppt.getSlides().get(slideIndex);
            for (int shapeIndex = 0; shapeIndex < slide.getShapes().size(); shapeIndex++) {
                HSLFShape shape = slide.getShapes().get(shapeIndex);
                if (shape instanceof HSLFTextShape) {
                    HSLFTextShape textShape = (HSLFTextShape) shape;
                    
                    String text = textShape.getText();
                    if (text != null && !text.trim().isEmpty()) {
                        Map<String, Object> position = new HashMap<>();
                        position.put("slideIndex", slideIndex);
                        position.put("shapeIndex", shapeIndex);
                        elements.add(new TextElement(text, "textShape", position));
                    }
                }
            }
        }
        return elements;
    }
    
    private void restorePPTTexts(HSLFSlideShow ppt, List<TextElement> elements, List<String> translatedTexts) {
        for (int i = 0; i < elements.size(); i++) {
            TextElement element = elements.get(i);
            int slideIndex = (Integer) element.position.get("slideIndex");
            int shapeIndex = (Integer) element.position.get("shapeIndex");
            
            HSLFSlide slide = ppt.getSlides().get(slideIndex);
            HSLFShape shape = slide.getShapes().get(shapeIndex);
            if (shape instanceof HSLFTextShape) {
                HSLFTextShape textShape = (HSLFTextShape) shape;
                textShape.setText(translatedTexts.get(i));
            }
        }
    }
}






