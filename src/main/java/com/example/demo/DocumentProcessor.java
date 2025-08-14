package com.example.demo;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hwpf.usermodel.*;
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

        System.out.println("翻译完成，开始写回文档");

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

        // 1.1 普通段落（跳过表格内段落）
        for (int pIdx = 0; pIdx < range.numParagraphs(); pIdx++) {
            Paragraph para = range.getParagraph(pIdx);
            if (para.isInTable()) continue;
            for (int rIdx = 0; rIdx < para.numCharacterRuns(); rIdx++) {
                CharacterRun run = para.getCharacterRun(rIdx);
                String txt = run.text();
                if (txt != null && !txt.trim().isEmpty()) {
                    Map<String,Object> pos = new HashMap<>();
                    pos.put("paragraphIndex", pIdx);
                    pos.put("runIndex", rIdx);
                    elements.add(new TextElement(txt, "run", pos));
                }
            }
        }

        // 1.2 表格：按表格→行→单元格→段落→Run 提取
        List<Table> tables = new ArrayList<>();
        TableIterator tit = new TableIterator(range);
        while (tit.hasNext()) {
            tables.add(tit.next());
        }
        for (int tIdx = 0; tIdx < tables.size(); tIdx++) {
            Table table = tables.get(tIdx);
            for (int rowIdx = 0; rowIdx < table.numRows(); rowIdx++) {
                TableRow row = table.getRow(rowIdx);
                for (int cellIdx = 0; cellIdx < row.numCells(); cellIdx++) {
                    TableCell cell = row.getCell(cellIdx);
                    // 一个单元格可能含多个段落
                    for (int cp = 0; cp < cell.numParagraphs(); cp++) {
                        Paragraph cellPara = cell.getParagraph(cp);
                        for (int cr = 0; cr < cellPara.numCharacterRuns(); cr++) {
                            CharacterRun run = cellPara.getCharacterRun(cr);
                            String txt = run.text();
                            if (txt != null) {
                                // 去掉控制字符
                                txt = txt.replaceAll("\\p{Cntrl}", "");
                            }
                            if (txt != null && !txt.trim().isEmpty()) {
                                Map<String,Object> pos = new HashMap<>();
                                pos.put("tableIndex", tIdx);
                                pos.put("rowIndex", rowIdx);
                                pos.put("cellIndex", cellIdx);
                                pos.put("cellParaIndex", cp);
                                pos.put("cellRunIndex", cr);
                                elements.add(new TextElement(txt, "tableCellRun", pos));
                            }
                        }
                    }
                }
            }
        }

        return elements;
    }

    /**
     * 恢复文档文本，包括普通段落 run 和表格 cell run，
     * 替换时生成绝对不与 old/new 重叠的安全 token
     */
    private void restoreWordTexts(HWPFDocument doc,List<TextElement> elements,List<String> translatedTexts) {
        Range docRange = doc.getRange();

        // 1. 缓存所有表格
        List<Table> tables = new ArrayList<>();
        TableIterator tit = new TableIterator(docRange);
        while (tit.hasNext()) {
            tables.add(tit.next());
        }

        // 2. 倒序替换，避免前面的替换影响后面的索引
        for (int idx = elements.size() - 1; idx >= 0; idx--) {
            TextElement el = elements.get(idx);
            String newRaw = translatedTexts.get(idx);
            if (newRaw == null) newRaw = "";
            newRaw = newRaw.replace("\n", "");

            String oldFull = null;
            CharacterRun run = null;
            boolean hasCR = false;

            if ("run".equals(el.type)) {
                Integer pI = (Integer) el.position.get("paragraphIndex");
                Integer rI = (Integer) el.position.get("runIndex");
                if (pI == null || rI == null) continue;
                if (pI < 0 || pI >= docRange.numParagraphs()) continue;
                Paragraph para = docRange.getParagraph(pI);
                if (rI < 0 || rI >= para.numCharacterRuns()) continue;

                run = para.getCharacterRun(rI);
                oldFull = run.text();
                if (oldFull == null) oldFull = "";
                hasCR = oldFull.endsWith("\r");

            } else if ("tableCellRun".equals(el.type)) {
                Integer tI  = (Integer) el.position.get("tableIndex");
                Integer rI  = (Integer) el.position.get("rowIndex");
                Integer cI  = (Integer) el.position.get("cellIndex");
                Integer cpI = (Integer) el.position.get("cellParaIndex");
                Integer crI = (Integer) el.position.get("cellRunIndex");
                if (tI == null || rI == null || cI == null || cpI == null || crI == null) continue;
                if (tI < 0 || tI >= tables.size()) continue;

                TableCell cell = tables.get(tI).getRow(rI).getCell(cI);
                Paragraph cellPara = cell.getParagraph(cpI);
                if (crI < 0 || crI >= cellPara.numCharacterRuns()) continue;
                run = cellPara.getCharacterRun(crI);

                oldFull = run.text();
                if (oldFull == null) oldFull = "";
                oldFull = oldFull.replaceAll("\\p{Cntrl}", ""); // 移除控制符
                hasCR = oldFull.endsWith("\r");
            }

            if (run != null && oldFull != null) {
                // === 生成绝对安全的 token ===
                String tokenCore = generateSafeToken(oldFull, newRaw);
                String token     = tokenCore + (hasCR ? "\r" : "");
                String newFull   = newRaw    + (hasCR ? "\r" : "");

                // 两步替换，避免 oldFull 与 newFull 部分重叠导致的死循环
                run.replaceText(oldFull, token);
                run.replaceText(token,   newFull);
            }
        }
    }

    /**
     * 生成一个在 oldFull 和 newCore 中都不含任一字符的唯一 token
     * 若全体字符都冲突，使用 Unicode 私有区（理论上极小概率）
     */
    private String generateSafeToken(String oldFull, String newCore) {
        Set<Character> forbidden = new HashSet<>();
        if (oldFull != null) {
            for (char c : oldFull.toCharArray()) forbidden.add(c);
        }
        if (newCore != null) {
            for (char c : newCore.toCharArray()) forbidden.add(c);
        }
        // 字符池：大写、小写、数字
        String base = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
        List<Character> allowed = new ArrayList<>();
        for (char c : base.toCharArray()) {
            if (!forbidden.contains(c)) allowed.add(c);
        }
        // 若允许集为空，则取 Unicode 私有区
        if (allowed.isEmpty()) {
            int code = 0xE000 + new Random().nextInt(0x1000);
            return new String(Character.toChars(code));
        }
        // 随机长度 8
        Random rnd = new Random();
        int length = 8;
        StringBuilder sb = new StringBuilder(length);
        for (int i = 0; i < length; i++) {
            sb.append(allowed.get(rnd.nextInt(allowed.size())));
        }
        return sb.toString();
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






