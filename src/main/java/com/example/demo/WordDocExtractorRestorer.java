package com.example.demo;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.*;

import java.lang.reflect.Method;
import java.util.*;

public class WordDocExtractorRestorer {

    public static class TextElement {
        public final String text;
        public final String type; // "run" | "tableCellRun" | "tbRun"
        public final Map<String, Object> position;
        public TextElement(String text, String type, Map<String, Object> position) {
            this.text = text; this.type = type; this.position = position;
        }
    }

    // ========= 抽取：正文/表格/文本框（跳过图片/对象锚点 run） =========
    public static List<TextElement> extractWordTexts(HWPFDocument doc) {
        List<TextElement> elements = new ArrayList<>();
        Range range = doc.getRange();
        PicturesTable pt = doc.getPicturesTable();

        // 非表格段落
        for (int pIdx = 0; pIdx < range.numParagraphs(); pIdx++) {
            Paragraph para = range.getParagraph(pIdx);
            if (para.isInTable()) continue;
            for (int rIdx = 0; rIdx < para.numCharacterRuns(); rIdx++) {
                CharacterRun run = para.getCharacterRun(rIdx);
                if (isPictureAnchor(run, pt)) continue;
                String clean = cleanForExtract(run.text());
                if (!clean.isEmpty()) {
                    Map<String,Object> pos = new HashMap<>();
                    pos.put("paragraphIndex", pIdx); pos.put("runIndex", rIdx);
                    elements.add(new TextElement(clean, "run", pos));
                }
            }
        }

        // 表格
        List<Table> tables = new ArrayList<>();
        TableIterator tit = new TableIterator(range);
        while (tit.hasNext()) tables.add(tit.next());
        for (int tIdx = 0; tIdx < tables.size(); tIdx++) {
            Table tbl = tables.get(tIdx);
            for (int r=0; r<tbl.numRows(); r++) {
                TableRow tr = tbl.getRow(r);
                for (int c=0; c<tr.numCells(); c++) {
                    TableCell cell = tr.getCell(c);
                    for (int p=0; p<cell.numParagraphs(); p++) {
                        Paragraph cp = cell.getParagraph(p);
                        for (int i=0; i<cp.numCharacterRuns(); i++) {
                            CharacterRun run = cp.getCharacterRun(i);
                            if (isPictureAnchor(run, pt)) continue;
                            String clean = cleanForExtract(run.text());
                            if (!clean.isEmpty()) {
                                Map<String,Object> pos = new HashMap<>();
                                pos.put("tableIndex", tIdx); pos.put("rowIndex", r);
                                pos.put("cellIndex", c); pos.put("cellParaIndex", p);
                                pos.put("cellRunIndex", i);
                                elements.add(new TextElement(clean, "tableCellRun", pos));
                            }
                        }
                    }
                }
            }
        }

        // 文本框
        Range tb = tryGetTextboxRange(doc);
        if (tb != null) {
            for (int p=0; p<tb.numParagraphs(); p++) {
                Paragraph para = tb.getParagraph(p);
                for (int r=0; r<para.numCharacterRuns(); r++) {
                    CharacterRun run = para.getCharacterRun(r);
                    if (isPictureAnchor(run, pt)) continue;
                    String clean = cleanForExtract(run.text());
                    if (!clean.isEmpty()) {
                        Map<String,Object> pos = new HashMap<>();
                        pos.put("tbParaIndex", p); pos.put("tbRunIndex", r);
                        elements.add(new TextElement(clean, "tbRun", pos));
                    }
                }
            }
        }

        return elements;
    }

    // ========= 写回：不修改内容，仅打印锚点前后对照 =========
    public static void restoreWordTexts(HWPFDocument doc,
                                        List<TextElement> elements,
                                        List<String> translatedTexts) {
        dumpAnchorsAll(doc, "before(no-change)");
        // 不做任何替换
        dumpAnchorsAll(doc, "after(no-change)");
    }

    // ===== 辅助 =====
    private static String cleanForExtract(String s) {
        if (s == null) return "";
        return s.replace("\r","").replaceAll("\\p{Cntrl}","").trim();
    }
    private static Range tryGetTextboxRange(HWPFDocument doc) {
        try { Range r = doc.getMainTextboxRange(); if (r != null) return r; } catch (Throwable ignore) {}
        try { Method m = HWPFDocument.class.getMethod("getTextboxesRange"); Object ret = m.invoke(doc);
              if (ret instanceof Range) return (Range) ret; } catch (Throwable ignore) {}
        return null;
    }
    private static boolean isPictureAnchor(CharacterRun cr, PicturesTable pt) {
        try {
            if (cr.isSpecialCharacter() || cr.isObj()) return true;
            if (pt != null && pt.hasPicture(cr)) return true;
            String t = cr.text();
            if (t != null) {
                for (int i=0;i<t.length();i++) {
                    char ch=t.charAt(i);
                    if (ch==0x01||ch==0x08||ch==0x13||ch==0x14||ch==0x15||ch==0xFFFC) return true;
                }
            }
        } catch (Throwable ignore) {}
        return false;
    }
    public static void dumpAnchorsAll(HWPFDocument doc, String tag) {
        PicturesTable pt = doc.getPicturesTable();
        Range rng = doc.getRange();
        System.out.println("["+tag+"] inline anchors:");
        for (int p=0;p<rng.numParagraphs();p++) {
            Paragraph para = rng.getParagraph(p);
            for (int r=0;r<para.numCharacterRuns();r++) {
                CharacterRun cr = para.getCharacterRun(r);
                if (isPictureAnchor(cr, pt)) {
                    System.out.printf("  p=%d r=%d start=%d end=%d%n",
                            p,r,cr.getStartOffset(),cr.getEndOffset());
                }
            }
        }
        try {
            java.lang.reflect.Field f = HWPFDocument.class.getDeclaredField("fspaTable");
            f.setAccessible(true);
            Object t = f.get(doc);
            if (t != null) {
                java.lang.reflect.Method m = t.getClass().getMethod("getShapes");
                @SuppressWarnings("unchecked")
                java.util.List<Object> shapes = (java.util.List<Object>) m.invoke(t);
                System.out.println("["+tag+"] floating anchors (FSPA cp):");
                for (Object s : shapes) {
                    int cp = (Integer)s.getClass().getMethod("getCp").invoke(s);
                    System.out.println("  cp="+cp);
                }
            }
        } catch (Throwable ignore) {
            System.out.println("["+tag+"] (no FSPA reflection)");
        }
    }
}
