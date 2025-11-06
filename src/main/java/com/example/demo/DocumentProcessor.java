package com.example.demo;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.hslf.usermodel.*;

import java.util.*;

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
    public XWPFDocument processWordDocument(XWPFDocument docx, String targetLang, String userPrompt) throws Exception {
        System.out.println("开始批量处理Word文档");

        List<WordDocxExtractorRestorer.TextElement> elements = WordDocxExtractorRestorer.extractWordTexts(docx);
        System.out.println("提取到 " + elements.size() + " 个文本元素");
        
        List<String> texts = new ArrayList<>();
        for (WordDocxExtractorRestorer.TextElement element : elements) {
            texts.add(element.text);
        }

        List<String> translatedTexts = translateService.batchTranslate(texts, targetLang, userPrompt);

        WordDocxExtractorRestorer.restoreWordTexts(docx, elements, translatedTexts);
        
        return docx;
    }
   
    // 2. Excel XLSX文档处理
    public XSSFWorkbook processExcelDocument(XSSFWorkbook workbook, String targetLang, String userPrompt) throws Exception {
        System.out.println("开始批量处理Excel文档");
        
        List<TextElement> elements = extractExcelTexts(workbook);
        System.out.println("提取到 " + elements.size() + " 个单元格文本");
        
        List<String> texts = new ArrayList<>();
        for (TextElement element : elements) {
            texts.add(element.text);
        }
        List<String> translatedTexts = translateService.batchTranslate(texts, targetLang, userPrompt);
        
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
    public XMLSlideShow processPowerPointPPTX(XMLSlideShow ppt, String targetLang, String userPrompt) throws Exception {
        System.out.println("开始批量处理PowerPoint PPTX文档");
        
        List<TextElement> elements = extractPPTXTexts(ppt);
        System.out.println("提取到 " + elements.size() + " 个文本元素");
        
        List<String> texts = new ArrayList<>();
        for (TextElement element : elements) {
            texts.add(element.text);
        }
        System.out.println(texts);
        List<String> translatedTexts = translateService.batchTranslate(texts, targetLang, userPrompt);
        
        restorePPTXTexts(ppt, elements, translatedTexts);
        return ppt;
    }
    
    private List<TextElement> extractPPTXTexts(XMLSlideShow ppt) {
        List<TextElement> elements = new ArrayList<>();

        for (int slideIndex = 0; slideIndex < ppt.getSlides().size(); slideIndex++) {
            XSLFSlide slide = ppt.getSlides().get(slideIndex);
            List<XSLFShape> shapes = slide.getShapes();

            for (int shapeIndex = 0; shapeIndex < shapes.size(); shapeIndex++) {
                // 递归遍历；top shape 的索引仍保留为 shapeIndex
                collectFromShape(shapes.get(shapeIndex), slideIndex, shapeIndex, new ArrayList<>(), elements);
            }
        }
        return elements;
    }

    private void collectFromShape(XSLFShape shape, int slideIndex, int topShapeIndex, List<Integer> path, List<TextElement> elements) {
        if (shape instanceof XSLFGroupShape group) {
            List<XSLFShape> children = group.getShapes();
            for (int i = 0; i < children.size(); i++) {
                ArrayList<Integer> nextPath = new ArrayList<>(path);
                nextPath.add(i);
                collectFromShape(children.get(i), slideIndex, topShapeIndex, nextPath, elements);
            }
            return;
        }

        if (shape instanceof XSLFTextShape textShape) {
            for (int paragraphIndex = 0; paragraphIndex < textShape.getTextParagraphs().size(); paragraphIndex++) {
                XSLFTextParagraph paragraph = textShape.getTextParagraphs().get(paragraphIndex);
                for (int runIndex = 0; runIndex < paragraph.getTextRuns().size(); runIndex++) {
                    XSLFTextRun run = paragraph.getTextRuns().get(runIndex);
                    String text = run.getRawText();
                    if (text != null && !text.trim().isEmpty()) {
                        Map<String, Object> position = new HashMap<>();
                        position.put("slideIndex", slideIndex);
                        position.put("shapeIndex", topShapeIndex); // 顶层 shape 的索引，保持兼容
                        position.put("paragraphIndex", paragraphIndex);
                        position.put("runIndex", runIndex);
                        if (!path.isEmpty()) position.put("shapePath", pathToString(path)); // 仅在分组中记录
                        elements.add(new TextElement(text, "textRun", position));
                    }
                }
            }
        } else if (shape instanceof XSLFTable table) {
            for (int rowIndex = 0; rowIndex < table.getRows().size(); rowIndex++) {
                XSLFTableRow row = table.getRows().get(rowIndex);
                for (int cellIndex = 0; cellIndex < row.getCells().size(); cellIndex++) {
                    XSLFTableCell cell = row.getCells().get(cellIndex);
                    String cellText = cell.getText();
                    if (cellText != null && !cellText.trim().isEmpty()) {
                        Map<String, Object> position = new HashMap<>();
                        position.put("slideIndex", slideIndex);
                        position.put("shapeIndex", topShapeIndex); // 顶层 shape 的索引，保持兼容
                        position.put("rowIndex", rowIndex);
                        position.put("cellIndex", cellIndex);
                        if (!path.isEmpty()) position.put("shapePath", pathToString(path));
                        elements.add(new TextElement(cellText, "tableCell", position));
                    }
                }
            }
        }
    }

    private String pathToString(List<Integer> path) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < path.size(); i++) {
            if (i > 0) sb.append('/');
            sb.append(path.get(i));
        }
        return sb.toString();
    }

    private void restorePPTXTexts(XMLSlideShow ppt, List<TextElement> elements, List<String> translatedTexts) {
        for (int i = 0; i < elements.size(); i++) {
            TextElement element = elements.get(i);
            String translatedText = translatedTexts.get(i);

            if ("textRun".equals(element.type)) {
                int slideIndex = (Integer) element.position.get("slideIndex");
                int shapeIndex = (Integer) element.position.get("shapeIndex");
                Integer paragraphIndex = (Integer) element.position.get("paragraphIndex");
                Integer runIndex = (Integer) element.position.get("runIndex");
                String shapePath = (String) element.position.get("shapePath"); // 可能为 null

                XSLFSlide slide = safeGet(ppt.getSlides(), slideIndex);
                if (slide == null) continue;

                XSLFShape shape = resolveShape(slide, shapeIndex, shapePath);
                if (!(shape instanceof XSLFTextShape)) continue;

                XSLFTextShape textShape = (XSLFTextShape) shape;
                XSLFTextParagraph paragraph = safeGet(textShape.getTextParagraphs(), paragraphIndex);
                if (paragraph == null) continue;
                XSLFTextRun run = safeGet(paragraph.getTextRuns(), runIndex);
                if (run == null) continue;

                run.setText(translatedText);

            } else if ("tableCell".equals(element.type)) {
                int slideIndex = (Integer) element.position.get("slideIndex");
                int shapeIndex = (Integer) element.position.get("shapeIndex");
                Integer rowIndex = (Integer) element.position.get("rowIndex");
                Integer cellIndex = (Integer) element.position.get("cellIndex");
                String shapePath = (String) element.position.get("shapePath"); // 可能为 null

                XSLFSlide slide = safeGet(ppt.getSlides(), slideIndex);
                if (slide == null) continue;

                XSLFShape shape = resolveShape(slide, shapeIndex, shapePath);
                if (!(shape instanceof XSLFTable)) continue;

                XSLFTable table = (XSLFTable) shape;
                XSLFTableRow row = safeGet(table.getRows(), rowIndex);
                if (row == null) continue;
                XSLFTableCell cell = safeGet(row.getCells(), cellIndex);
                if (cell == null) continue;

                cell.clearText();
                cell.setText(translatedText);
            }
        }
    }

    private XSLFShape resolveShape(XSLFSlide slide, int topShapeIndex, String shapePath) {
        XSLFShape current = safeGet(slide.getShapes(), topShapeIndex);
        if (current == null) return null;
        if (shapePath == null || shapePath.isEmpty()) return current;

        String[] parts = shapePath.split("/");
        for (String p : parts) {
            if (p.isEmpty()) continue;
            int idx;
            try { idx = Integer.parseInt(p); } catch (NumberFormatException e) { return null; }
            if (!(current instanceof XSLFGroupShape)) return null;
            XSLFGroupShape group = (XSLFGroupShape) current;
            current = safeGet(group.getShapes(), idx);
            if (current == null) return null;
        }
        return current;
    }

    private static <T> T safeGet(List<T> list, Integer idx) {
        if (list == null || idx == null) return null;
        return (idx >= 0 && idx < list.size()) ? list.get(idx) : null;
    }
    
    // 4. Word DOC处理
    public HWPFDocument processWordDOC(HWPFDocument doc, String targetLang, String userPrompt) throws Exception {
        System.out.println("开始批量处理 .doc 文档");
        List<WordDocExtractorRestorer.TextElement> elements = WordDocExtractorRestorer.extractWordTexts(doc);
        System.out.println("提取到 " + elements.size() + " 个文本元素");
        // 提取原文去翻译
        List<String> originals = new ArrayList<>();
        for (WordDocExtractorRestorer.TextElement el : elements) {
            originals.add(el.text);
        }
        // 批量翻译
        List<String> translated = translateService.batchTranslate(originals, targetLang, userPrompt);
        System.out.println("翻译完成，开始写回文档");
        // 写回翻译结果
        WordDocExtractorRestorer.restoreWordTexts(doc, elements, translated);
        return doc;
    }

    // 5. Excel XLS处理
    public HSSFWorkbook processExcelXLS(HSSFWorkbook workbook, String targetLang, String userPrompt) throws Exception {
        System.out.println("开始批量处理Excel XLS文档");
        
        List<TextElement> elements = extractXLSTexts(workbook);
        System.out.println("提取到 " + elements.size() + " 个单元格文本");
        
        List<String> texts = new ArrayList<>();
        for (TextElement element : elements) {
            texts.add(element.text);
        }
        
        if (!texts.isEmpty()) {
            List<String> translatedTexts = translateService.batchTranslate(texts, targetLang, userPrompt);
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
    public HSLFSlideShow processPowerPointPPT(HSLFSlideShow ppt, String targetLang, String userPrompt) throws Exception {
        System.out.println("开始批量处理PowerPoint PPT文档");
        
        List<TextElement> elements = extractPPTTexts(ppt);
        System.out.println("提取到 " + elements.size() + " 个文本元素");
        
        List<String> texts = new ArrayList<>();
        for (TextElement element : elements) {
            texts.add(element.text);
        }
        List<String> translatedTexts = translateService.batchTranslate(texts, targetLang, userPrompt);
        
        restorePPTTexts(ppt, elements, translatedTexts);
        return ppt;
    }
    
    private List<TextElement> extractPPTTexts(HSLFSlideShow ppt) {
        List<TextElement> elements = new ArrayList<>();

        for (int slideIndex = 0; slideIndex < ppt.getSlides().size(); slideIndex++) {
            HSLFSlide slide = ppt.getSlides().get(slideIndex);
            List<HSLFShape> shapes = slide.getShapes();

            for (int shapeIndex = 0; shapeIndex < shapes.size(); shapeIndex++) {
                collectFromHslfShape(shapes.get(shapeIndex), slideIndex, shapeIndex, new ArrayList<>(), elements);
            }
        }
        return elements;
    }

    private void collectFromHslfShape(HSLFShape shape, int slideIndex, int topShapeIndex, List<Integer> path, List<TextElement> out) {
        if (shape instanceof HSLFGroupShape) {
            HSLFGroupShape group = (HSLFGroupShape) shape;
            List<HSLFShape> children = group.getShapes();
            for (int i = 0; i < children.size(); i++) {
                ArrayList<Integer> next = new ArrayList<>(path);
                next.add(i);
                collectFromHslfShape(children.get(i), slideIndex, topShapeIndex, next, out);
            }
            return;
        }

        if (shape instanceof HSLFTextShape) {
            HSLFTextShape textShape = (HSLFTextShape) shape;
            String text = textShape.getText();
            if (text != null && !text.trim().isEmpty()) {
                Map<String, Object> position = new HashMap<>();
                position.put("slideIndex", slideIndex);
                position.put("shapeIndex", topShapeIndex);       // 仍记录顶层 shape 索引，保持兼容
                if (!path.isEmpty()) position.put("shapePath", pathToString(path));
                out.add(new TextElement(text, "textShape", position));
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
                            position.put("shapeIndex", topShapeIndex);
                            position.put("rowIndex", rowIndex);
                            position.put("cellIndex", cellIndex);
                            if (!path.isEmpty()) position.put("shapePath", pathToString(path));
                            out.add(new TextElement(cellText, "tableCell", position));
                        }
                    }
                }
            }
        }
    }

    private void restorePPTTexts(HSLFSlideShow ppt, List<TextElement> elements, List<String> translatedTexts) {
        for (int i = 0; i < elements.size(); i++) {
            TextElement element = elements.get(i);
            String translatedText = translatedTexts.get(i);

            if ("textShape".equals(element.type)) {
                Integer slideIndex = (Integer) element.position.get("slideIndex");
                Integer shapeIndex = (Integer) element.position.get("shapeIndex");
                String shapePath = (String) element.position.get("shapePath"); // 可能为 null

                HSLFSlide slide = safeGet(ppt.getSlides(), slideIndex);
                if (slide == null) continue;

                HSLFShape shape = resolveHslfShape(slide, shapeIndex, shapePath);
                if (!(shape instanceof HSLFTextShape)) continue;

                ((HSLFTextShape) shape).setText(translatedText);

            } else if ("tableCell".equals(element.type)) {
                Integer slideIndex = (Integer) element.position.get("slideIndex");
                Integer shapeIndex = (Integer) element.position.get("shapeIndex");
                Integer rowIndex = (Integer) element.position.get("rowIndex");
                Integer cellIndex = (Integer) element.position.get("cellIndex");
                String shapePath = (String) element.position.get("shapePath"); // 可能为 null

                HSLFSlide slide = safeGet(ppt.getSlides(), slideIndex);
                if (slide == null) continue;

                HSLFShape shape = resolveHslfShape(slide, shapeIndex, shapePath);
                if (!(shape instanceof HSLFTable)) continue;

                HSLFTable table = (HSLFTable) shape;
                HSLFTableCell cell = (rowIndex != null && cellIndex != null) ? table.getCell(rowIndex, cellIndex) : null;
                if (cell != null) {
                    cell.setText(translatedText);
                }
            }
        }
    }

    private HSLFShape resolveHslfShape(HSLFSlide slide, Integer topShapeIndex, String shapePath) {
        HSLFShape current = safeGet(slide.getShapes(), topShapeIndex);
        if (current == null) return null;
        if (shapePath == null || shapePath.isEmpty()) return current;

        String[] parts = shapePath.split("/");
        for (String p : parts) {
            if (!(current instanceof HSLFGroupShape)) return null;
            int idx;
            try { idx = Integer.parseInt(p); } catch (NumberFormatException e) { return null; }
            current = safeGet(((HSLFGroupShape) current).getShapes(), idx);
            if (current == null) return null;
        }
        return current;
    }

}