package com.example.demo;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
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
        
        // 处理段落
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
        
        // 处理表格
        for (int tableIndex = 0; tableIndex < doc.getTables().size(); tableIndex++) {
            XWPFTable table = doc.getTables().get(tableIndex);
            for (int rowIndex = 0; rowIndex < table.getRows().size(); rowIndex++) {
                XWPFTableRow row = table.getRows().get(rowIndex);
                for (int cellIndex = 0; cellIndex < row.getTableCells().size(); cellIndex++) {
                    XWPFTableCell cell = row.getTableCells().get(cellIndex);
                    String cellText = cell.getText();
                    if (cellText != null && !cellText.trim().isEmpty()) {
                        Map<String, Object> position = new HashMap<>();
                        position.put("tableIndex", tableIndex);
                        position.put("rowIndex", rowIndex);
                        position.put("cellIndex", cellIndex);
                        elements.add(new TextElement(cellText, "tableCell", position));
                    }
                }
            }
        }
        
        return elements;
    }
    
    private void restoreWordTexts(XWPFDocument doc, List<TextElement> elements, List<String> translatedTexts) {
        for (int i = 0; i < elements.size(); i++) {
            TextElement element = elements.get(i);
            String translatedText = translatedTexts.get(i);
            
            if ("run".equals(element.type)) {
                int paragraphIndex = (Integer) element.position.get("paragraphIndex");
                int runIndex = (Integer) element.position.get("runIndex");
                
                XWPFRun run = doc.getParagraphs().get(paragraphIndex).getRuns().get(runIndex);
                run.setText(translatedText, 0);
                
            } else if ("tableCell".equals(element.type)) {
                int tableIndex = (Integer) element.position.get("tableIndex");
                int rowIndex = (Integer) element.position.get("rowIndex");
                int cellIndex = (Integer) element.position.get("cellIndex");
                
                XWPFTable table = doc.getTables().get(tableIndex);
                XWPFTableCell cell = table.getRows().get(rowIndex).getTableCells().get(cellIndex);
                
                // 清除原有内容并设置新文本
                cell.removeParagraph(0);
                XWPFParagraph newParagraph = cell.addParagraph();
                XWPFRun newRun = newParagraph.createRun();
                newRun.setText(translatedText);
            }
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
                } else if (shape instanceof XSLFTable) {
                    XSLFTable table = (XSLFTable) shape;
                    
                    for (int rowIndex = 0; rowIndex < table.getRows().size(); rowIndex++) {
                        XSLFTableRow row = table.getRows().get(rowIndex);
                        for (int cellIndex = 0; cellIndex < row.getCells().size(); cellIndex++) {
                            XSLFTableCell cell = row.getCells().get(cellIndex);
                            String cellText = cell.getText();
                            if (cellText != null && !cellText.trim().isEmpty()) {
                                Map<String, Object> position = new HashMap<>();
                                position.put("slideIndex", slideIndex);
                                position.put("shapeIndex", shapeIndex);
                                position.put("rowIndex", rowIndex);
                                position.put("cellIndex", cellIndex);
                                elements.add(new TextElement(cellText, "tableCell", position));
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
            String translatedText = translatedTexts.get(i);
            
            if ("textRun".equals(element.type)) {
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
                    
                    run.setText(translatedText);
                }
            } else if ("tableCell".equals(element.type)) {
                int slideIndex = (Integer) element.position.get("slideIndex");
                int shapeIndex = (Integer) element.position.get("shapeIndex");
                int rowIndex = (Integer) element.position.get("rowIndex");
                int cellIndex = (Integer) element.position.get("cellIndex");
                
                XSLFSlide slide = ppt.getSlides().get(slideIndex);
                XSLFShape shape = slide.getShapes().get(shapeIndex);
                if (shape instanceof XSLFTable) {
                    XSLFTable table = (XSLFTable) shape;
                    XSLFTableCell cell = table.getRows().get(rowIndex).getCells().get(cellIndex);
                    cell.clearText();
                    cell.setText(translatedText);
                }
            }
        }
    }
    
    // 4. Word DOC处理
    public HWPFDocument processWordDOC(HWPFDocument doc) throws Exception {
        System.out.println("开始批量处理 .doc 文档");

        List<TextElement> elements = extractWordTexts(doc);
        System.out.println("提取到 " + elements.size() + " 个文本元素");

        // 提取原文去翻译
        List<String> originals = new ArrayList<>();
        for (TextElement el : elements) {
            originals.add(el.text);
        }
        // 批量翻译
        List<String> translated = translateService.batchTranslate(originals);

        // 写回翻译结果
        restoreWordTexts(doc, elements, translated);

        return doc;
    }

    /**
     * 遍历 Range 中的所有段落和 CharacterRun，
     * 把非空文本以 TextElement 形式保存。
     */
    private List<TextElement> extractWordTexts(HWPFDocument doc) {
        List<TextElement> elements = new ArrayList<>();
        Range range = doc.getRange();
        int numParas = range.numParagraphs();

        for (int pIdx = 0; pIdx < numParas; pIdx++) {
            Paragraph para = range.getParagraph(pIdx);
            int numRuns = para.numCharacterRuns();
            for (int rIdx = 0; rIdx < numRuns; rIdx++) {
                CharacterRun run = para.getCharacterRun(rIdx);
                String txt = run.text();
                if (txt != null && !txt.trim().isEmpty()) {
                    Map<String, Object> pos = new HashMap<>();
                    pos.put("paragraphIndex", pIdx);
                    pos.put("runIndex", rIdx);
                    // type 标记为 "run"
                    elements.add(new TextElement(txt, "run", pos));
                }
            }
        }
        return elements;
    }

    /**
     * 遍历 Range 中的所有段落和 CharacterRun，
     * 将原文替换为唯一占位符，
     * 然后再将占位符替换为翻译后的文本，
     * 这样可以避免直接替换时可能出现的无限循环问题。  
     */
    private void restoreWordTexts(HWPFDocument doc,
                                    List<TextElement> elements,
                                    List<String> translatedTexts) {

        Range range = doc.getRange();

        // 步骤 1: 将原文替换为唯一占位符
        Map<String, String> placeholderMap = new HashMap<>();

        for (int i = 0; i < elements.size(); i++) {
            TextElement el = elements.get(i);
            String oldText = el.text;

            // 跳过空文本
            if (oldText == null || oldText.trim().isEmpty()) continue;

            oldText = oldText.replace("\r", "");  // remove CRs
            oldText = oldText.trim();

            String placeholder = "<<<REPLACE_" + i + ">>>";
            placeholderMap.put(placeholder, translatedTexts.get(i));

            System.out.println("步骤 1 - 替换:");
            System.out.println("  原文: \"" + oldText + "\"");
            System.out.println("  占位符: \"" + placeholder + "\"");

            // 使用 Range 的 replaceText 方法替换原文为占位符
            range.replaceText(oldText, placeholder);
        }

        // 步骤 2: 将占位符替换为翻译后的文本
        for (Map.Entry<String, String> entry : placeholderMap.entrySet()) {
            String placeholder = entry.getKey();
            String translated = entry.getValue();

            System.out.println("步骤 2 - 替换占位符:");
            System.out.println("  占位符: \"" + placeholder + "\"");
            System.out.println("  翻译后: \"" + translated + "\"");

            range.replaceText(placeholder, translated);
        }

        System.out.println("文本恢复完成，翻译结果已写入文档。");
    }
    
    // 5. Excel XLS处理
    public HSSFWorkbook processExcelXLS(HSSFWorkbook workbook) throws Exception {
        System.out.println("开始批量处理Excel XLS文档");
        
        List<TextElement> elements = extractXLSTexts(workbook);
        System.out.println("提取到 " + elements.size() + " 个单元格文本");
        
        List<String> texts = new ArrayList<>();
        for (TextElement element : elements) {
            texts.add(element.text);
        }
        
        if (!texts.isEmpty()) {
            List<String> translatedTexts = translateService.batchTranslate(texts);
            restoreXLSTexts(workbook, elements, translatedTexts);
        }
        
        return workbook;
    }

    private List<TextElement> extractXLSTexts(HSSFWorkbook workbook) {
        List<TextElement> elements = new ArrayList<>();
        
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            HSSFSheet sheet = workbook.getSheetAt(sheetIndex);
            for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                HSSFRow row = sheet.getRow(rowNum);
                if (row != null) {
                    for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
                        HSSFCell cell = row.getCell(cellNum);
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

    private void restoreXLSTexts(HSSFWorkbook workbook, List<TextElement> elements, List<String> translatedTexts) {
        for (int i = 0; i < elements.size(); i++) {
            TextElement element = elements.get(i);
            int sheetIndex = (Integer) element.position.get("sheetIndex");
            int rowIndex = (Integer) element.position.get("rowIndex");
            int cellIndex = (Integer) element.position.get("cellIndex");
            
            HSSFCell cell = workbook.getSheetAt(sheetIndex).getRow(rowIndex).getCell(cellIndex);
            cell.setCellValue(translatedTexts.get(i));
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
                } else if (shape instanceof HSLFTable) {
                    HSLFTable table = (HSLFTable) shape;
                    
                    for (int rowIndex = 0; rowIndex < table.getNumberOfRows(); rowIndex++) {
                        for (int cellIndex = 0; cellIndex < table.getNumberOfColumns(); cellIndex++) {
                            HSLFTableCell cell = table.getCell(rowIndex, cellIndex);
                            if (cell != null) {
                                String cellText = cell.getText();
                                if (cellText != null && !cellText.trim().isEmpty()) {
                                    Map<String, Object> position = new HashMap<>();
                                    position.put("slideIndex", slideIndex);
                                    position.put("shapeIndex", shapeIndex);
                                    position.put("rowIndex", rowIndex);
                                    position.put("cellIndex", cellIndex);
                                    elements.add(new TextElement(cellText, "tableCell", position));
                                }
                            }
                        }
                    }
                }
            }
        }
        return elements;
    }
    
    private void restorePPTTexts(HSLFSlideShow ppt, List<TextElement> elements, List<String> translatedTexts) {
        for (int i = 0; i < elements.size(); i++) {
            TextElement element = elements.get(i);
            String translatedText = translatedTexts.get(i);
            
            if ("textShape".equals(element.type)) {
                int slideIndex = (Integer) element.position.get("slideIndex");
                int shapeIndex = (Integer) element.position.get("shapeIndex");
                
                HSLFSlide slide = ppt.getSlides().get(slideIndex);
                HSLFShape shape = slide.getShapes().get(shapeIndex);
                if (shape instanceof HSLFTextShape) {
                    HSLFTextShape textShape = (HSLFTextShape) shape;
                    textShape.setText(translatedText);
                }
            } else if ("tableCell".equals(element.type)) {
                int slideIndex = (Integer) element.position.get("slideIndex");
                int shapeIndex = (Integer) element.position.get("shapeIndex");
                int rowIndex = (Integer) element.position.get("rowIndex");
                int cellIndex = (Integer) element.position.get("cellIndex");
                
                HSLFSlide slide = ppt.getSlides().get(slideIndex);
                HSLFShape shape = slide.getShapes().get(shapeIndex);
                if (shape instanceof HSLFTable) {
                    HSLFTable table = (HSLFTable) shape;
                    HSLFTableCell cell = table.getCell(rowIndex, cellIndex);
                    if (cell != null) {
                        cell.setText(translatedText);
                    }
                }
            }
        }
    }
}






