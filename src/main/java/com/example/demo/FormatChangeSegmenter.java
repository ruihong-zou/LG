// File: src/main/java/com/example/demo/FormatChangeSegmenter.java
package com.example.demo;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

import javax.xml.namespace.QName;
import java.util.*;

/** 对段落进行“仅因格式变化分段”，支持宽松合并策略与多容器遍历 */
public final class FormatChangeSegmenter {
    private FormatChangeSegmenter() {}
    private static final String NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /** 段片段 */
    public static final class Segment {
        public final XWPFParagraph paragraph;
        public final int runStartIdx;
        public final int runEndIdx;
        public final StyleKey styleKey;
        public final String text;

        public Segment(XWPFParagraph p, int s, int e, StyleKey k, String text) {
            this.paragraph = p; this.runStartIdx = s; this.runEndIdx = e; this.styleKey = k; this.text = text;
        }
    }

    /** 样式指纹（公开 API + 轻量 XML 读取） */
    public static final class StyleKey {
        public final String fontFamily;
        public final Integer fontSizePt;
        public final boolean bold, italic, strike;
        public final String underline, colorHex, vertAlign;
        public final boolean inHyperlink, inField;
        public final Integer charSpacingHundPt;
        public final String highlightVal;

        public StyleKey(String fontFamily, Integer fontSizePt, boolean bold, boolean italic, boolean strike,
                        String underline, String colorHex, String vertAlign, boolean inHyperlink, boolean inField,
                        Integer charSpacingHundPt, String highlightVal) {
            this.fontFamily = nz(fontFamily);
            this.fontSizePt = (fontSizePt != null && fontSizePt >= 0) ? fontSizePt : null;
            this.bold = bold; this.italic = italic; this.strike = strike;
            this.underline = nz(underline);
            this.colorHex = normHex(colorHex);
            this.vertAlign = nz(vertAlign);
            this.inHyperlink = inHyperlink; this.inField = inField;
            this.charSpacingHundPt = charSpacingHundPt;
            this.highlightVal = nz(highlightVal);
        }
        private static String nz(String s){ return (s==null||s.isEmpty())?null:s; }
        private static String normHex(String s){
            if (s==null||s.isEmpty()) return null;
            String t = s.startsWith("#")?s.substring(1):s;
            if (t.length()==8) t=t.substring(2);
            return t.length()==6 ? t.toLowerCase(Locale.ROOT) : t.toLowerCase(Locale.ROOT);
        }
    }

    /** 遍历文档各容器并分段 */
    public static List<Segment> segmentDocument(XWPFDocument doc) {
        List<Segment> out = new ArrayList<>();
        segmentBody(doc, out);
        for (XWPFHeader h : doc.getHeaderList()) segmentBody(h, out);
        for (XWPFFooter f : doc.getFooterList()) segmentBody(f, out);
        for (XWPFFootnote fn : doc.getFootnotes()) segmentBody(fn, out);
        return out;
    }

    private static void segmentBody(IBody body, List<Segment> out) {
        for (IBodyElement be : body.getBodyElements()) {
            if (be instanceof XWPFParagraph) {
                out.addAll(segmentParagraph((XWPFParagraph) be, MergePolicy.loose()));
            } else if (be instanceof XWPFTable) {
                segmentTable((XWPFTable) be, out);
            } else if (be instanceof XWPFSDT) {
                ISDTContent content = ((XWPFSDT) be).getContent();
                boolean handled = false;
                try {
                    var m = content.getClass().getMethod("getBodyElements");
                    @SuppressWarnings("unchecked")
                    List<IBodyElement> bes = (List<IBodyElement>) m.invoke(content);
                    for (IBodyElement inner : bes) {
                        if (inner instanceof XWPFParagraph) out.addAll(segmentParagraph((XWPFParagraph) inner, MergePolicy.loose()));
                        else if (inner instanceof XWPFTable) segmentTable((XWPFTable) inner, out);
                    }
                    handled = true;
                } catch (Throwable ignored) {}
                if (!handled) {
                    try {
                        var mp = content.getClass().getMethod("getParagraphs");
                        @SuppressWarnings("unchecked")
                        List<XWPFParagraph> ps = (List<XWPFParagraph>) mp.invoke(content);
                        for (XWPFParagraph p : ps) out.addAll(segmentParagraph(p, MergePolicy.loose()));
                    } catch (Throwable ignored) {}
                    try {
                        var mt = content.getClass().getMethod("getTables");
                        @SuppressWarnings("unchecked")
                        List<XWPFTable> ts = (List<XWPFTable>) mt.invoke(content);
                        for (XWPFTable t : ts) segmentTable(t, out);
                    } catch (Throwable ignored) {}
                }
            }
        }
    }

    private static void segmentTable(XWPFTable t, List<Segment> out) {
        for (XWPFTableRow row : t.getRows())
            for (XWPFTableCell cell : row.getTableCells())
                segmentBody(cell, out);
    }

    /** 段落分段（带 MergePolicy 的软等价） */
    public static List<Segment> segmentParagraph(XWPFParagraph p, MergePolicy policy) {
        List<Segment> out = new ArrayList<>();
        List<XWPFRun> runs = p.getRuns();
        if (runs == null || runs.isEmpty()) return out;

        int segStart = 0;
        StyleKey segKey = keyOf(runs.get(0));
        StringBuilder segText = new StringBuilder();

        for (int i = 0; i < runs.size(); i++) {
            XWPFRun r = runs.get(i);
            boolean boundary = (i == 0) ? false : hardBoundary(r, runs.get(i - 1));
            StyleKey k = keyOf(r);
            if (i > 0 && (boundary || !softEqual(segKey, k, policy))) {
                out.add(new Segment(p, segStart, i - 1, segKey, segText.toString()));
                segStart = i; segKey = k; segText.setLength(0);
            }
            appendRunTextPreserve(r, segText);
        }
        out.add(new Segment(p, segStart, runs.size() - 1, segKey, segText.toString()));
        return out;
    }

    public static StyleKey keyOf(XWPFRun r) {
        String family = r.getFontFamily();
        Integer sizePt = fontSizePtOf(r);
        boolean bold = r.isBold();
        boolean italic = r.isItalic();
        boolean strike = r.isStrikeThrough();
        String underline = (r.getUnderline()!=null) ? r.getUnderline().name() : null;
        String colorHex = r.getColor();
        String vertAlign = vertAlignOf(r);
        boolean inHyperlink = (r instanceof XWPFHyperlinkRun);
        boolean inField = hasFieldMark(r.getCTR());
        Integer charSp = charSpacingHundPtOf(r);
        String highlight = highlightOf(r);
        return new StyleKey(family, sizePt, bold, italic, strike, underline, colorHex, vertAlign,
                inHyperlink, inField, charSp, highlight);
    }

    private static boolean softEqual(StyleKey a, StyleKey b, MergePolicy p) {
        if (a.inHyperlink != b.inHyperlink) return false;
        if (a.inField != b.inField) return false;
        if (!Objects.equals(nz(a.vertAlign), nz(b.vertAlign))) return false;

        if (a.fontSizePt != null && b.fontSizePt != null) {
            if (Math.abs(a.fontSizePt - b.fontSizePt) > p.fontSizeTolerancePt) return false;
        }
        if (a.charSpacingHundPt != null && b.charSpacingHundPt != null) {
            if (Math.abs(a.charSpacingHundPt - b.charSpacingHundPt) > p.charSpacingTolerance) return false;
        }

        if (!p.hardOnFontFamilyDiff) {
            if (p.ignoreFontFamilyNullVsExplicit) {
                if (a.fontFamily==null || b.fontFamily==null) { /* pass */ }
                else if (!a.fontFamily.equals(b.fontFamily)) return false;
            } else if (!Objects.equals(a.fontFamily, b.fontFamily)) return false;
        } else if (!Objects.equals(a.fontFamily, b.fontFamily)) return false;

        if (!p.ignoreAllColorDiff) {
            String ac = normColorForCompare(a.colorHex, p);
            String bc = normColorForCompare(b.colorHex, p);
            if (!Objects.equals(ac, bc)) return false;
        }

        if (!p.ignoreHighlightDiff) {
            if (!Objects.equals(a.highlightVal, b.highlightVal)) return false;
        }

        if (p.hardOnBoldDiff && a.bold != b.bold) return false;
        if (p.hardOnItalicDiff && a.italic != b.italic) return false;
        if (p.hardOnUnderlineDiff && !Objects.equals(a.underline, b.underline)) return false;
        if (p.hardOnStrikeDiff && a.strike != b.strike) return false;

        return true;
    }

    private static String normColorForCompare(String hex, MergePolicy p) {
        if (hex == null) return p.ignoreColorAutoVsNull ? "auto" : null;
        String h = hex.startsWith("#") ? hex.substring(1) : hex;
        if (h.equalsIgnoreCase("auto")) return p.ignoreColorAutoVsNull ? "auto" : "auto!";
        if (h.length()==8) h=h.substring(2);
        return h.length()==6 ? h.toLowerCase(Locale.ROOT) : h.toLowerCase(Locale.ROOT);
    }

    private static boolean hardBoundary(XWPFRun cur, XWPFRun prev) {
        if ((cur instanceof XWPFHyperlinkRun)!=(prev instanceof XWPFHyperlinkRun)) return true;
        boolean curFld = hasFieldMark(cur.getCTR());
        boolean prevFld = hasFieldMark(prev.getCTR());
        return curFld != prevFld;
    }

    private static Integer fontSizePtOf(XWPFRun r) {
        CTR ctr = r.getCTR();
        if (ctr != null) {
            try (XmlCursor c = ctr.newCursor()) {
                c.selectPath("declare namespace w='" + NS_W + "' .//w:sz");
                if (c.toNextSelection()) {
                    String half = c.getAttributeText(new QName(NS_W, "val"));
                    if (half != null && !half.isEmpty()) {
                        try {
                            int halfPt = Integer.parseInt(half);
                            return Math.max(0, (int) Math.round(halfPt / 2.0));
                        } catch (NumberFormatException ignore) {}
                    }
                }
            }
        }
        try {
            @SuppressWarnings("deprecation")
            int v = r.getFontSize();
            return (v > 0) ? v : null;
        } catch (Throwable ignore) { return null; }
    }

    private static String vertAlignOf(XWPFRun r) {
        CTR ctr = r.getCTR(); if (ctr == null) return null;
        try (XmlCursor c = ctr.newCursor()) {
            c.selectPath("declare namespace w='" + NS_W + "' .//w:vertAlign");
            if (c.toNextSelection()) {
                String val = c.getAttributeText(new QName(NS_W, "val"));
                return (val == null || val.isEmpty()) ? null : val;
            }
        }
        return null;
    }

    private static Integer charSpacingHundPtOf(XWPFRun r) {
        CTR ctr = r.getCTR(); if (ctr == null) return null;
        try (XmlCursor c = ctr.newCursor()) {
            c.selectPath("declare namespace w='" + NS_W + "' .//w:rPr/w:spacing");
            if (c.toNextSelection()) {
                String v = c.getAttributeText(new QName(NS_W,"val"));
                if (v != null && !v.isEmpty()) {
                    try { return Integer.parseInt(v); } catch (NumberFormatException ignore) {}
                }
            }
        }
        return null;
    }

    private static String highlightOf(XWPFRun r) {
        CTR ctr = r.getCTR(); if (ctr == null) return null;
        try (XmlCursor c = ctr.newCursor()) {
            c.selectPath("declare namespace w='" + NS_W + "' .//w:rPr/w:highlight");
            if (c.toNextSelection()) {
                String v = c.getAttributeText(new QName(NS_W,"val"));
                return (v == null || v.isEmpty()) ? null : v;
            }
        }
        return null;
    }

    private static boolean hasFieldMark(CTR ctr) {
        if (ctr == null) return false;
        if (ctr.sizeOfFldCharArray() > 0) return true;
        if (ctr.sizeOfInstrTextArray() > 0) return true;
        return false;
    }

    private static void appendRunTextPreserve(XWPFRun r, StringBuilder sb) {
        CTR ctr = r.getCTR(); if (ctr == null) return;
        for (CTText t : ctr.getTList()) { if (t != null) {
            String v = t.getStringValue(); if (v != null) sb.append(v);
        }}
        for (int i = 0; i < ctr.sizeOfBrArray(); i++) sb.append('\n');
        for (int i = 0; i < ctr.sizeOfTabArray(); i++) sb.append('\t');
    }

    private static String nz(String s){ return (s==null||s.isEmpty())?null:s; }
}
