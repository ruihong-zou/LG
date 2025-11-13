package com.example.demo;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.apache.xmlbeans.XmlCursor;

import javax.xml.namespace.QName;
import java.util.*;

public final class FormatChangeSegmenter {

    private FormatChangeSegmenter() {}

    private static final String NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // ---------- 数据模型 ----------
    public static final class Segment {
        public final XWPFParagraph paragraph;
        public final int runStartIdx;
        public final int runEndIdx;
        public final StyleKey styleKey;
        public final String text;

        public Segment(XWPFParagraph p, int s, int e, StyleKey k, String text) {
            this.paragraph = p; this.runStartIdx = s; this.runEndIdx = e; this.styleKey = k; this.text = text;
        }

        @Override public String toString() {
            return "Segment{runs=" + runStartIdx + ".." + runEndIdx +
                    ", style=" + styleKey + ", text=\"" +
                    text.replace("\n","\\n").replace("\t","\\t") + "\"}";
        }
    }

    /** 精简样式指纹（仅用公开 API + 光标扫描，兼容老版本 POI） */
    public static final class StyleKey {
        public final String fontFamily;        // 来自 run.getFontFamily()
        public final Integer fontSizePt;       // 来自 <w:sz w:val>（half-points→pt），取不到再回退 getFontSize()
        public final boolean bold;
        public final boolean italic;
        public final boolean strike;
        public final String underline;         // run.getUnderline().name()
        public final String colorHex;          // run.getColor() 规范化
        public final String vertAlign;         // 来自 <w:vertAlign w:val>
        public final boolean inHyperlink;      // run instanceof XWPFHyperlinkRun
        public final boolean inField;          // CTR 中是否含域标记/指令

        public StyleKey(String fontFamily, Integer fontSizePt, boolean bold, boolean italic, boolean strike,
                        String underline, String colorHex, String vertAlign, boolean inHyperlink, boolean inField) {
            this.fontFamily = nz(fontFamily);
            this.fontSizePt = (fontSizePt != null && fontSizePt >= 0) ? fontSizePt : null;
            this.bold = bold;
            this.italic = italic;
            this.strike = strike;
            this.underline = nz(underline);
            this.colorHex = normHex(colorHex);
            this.vertAlign = nz(vertAlign);
            this.inHyperlink = inHyperlink;
            this.inField = inField;
        }

        private static String nz(String s) { return (s == null || s.isEmpty()) ? null : s; }
        private static String normHex(String s) {
            if (s == null || s.isEmpty()) return null;
            String t = s.startsWith("#") ? s.substring(1) : s;
            if (t.length() == 8) t = t.substring(2);
            if (t.length() != 6) return null;
            return t.toLowerCase(Locale.ROOT);
        }

        @Override public boolean equals(Object o) {
            if (this == o) return true;
            if (!(o instanceof StyleKey)) return false;
            StyleKey k = (StyleKey) o;
            return bold == k.bold && italic == k.italic && strike == k.strike &&
                   inHyperlink == k.inHyperlink && inField == k.inField &&
                   Objects.equals(fontFamily, k.fontFamily) &&
                   Objects.equals(fontSizePt, k.fontSizePt) &&
                   Objects.equals(underline, k.underline) &&
                   Objects.equals(colorHex, k.colorHex) &&
                   Objects.equals(vertAlign, k.vertAlign);
        }

        @Override public int hashCode() {
            return Objects.hash(fontFamily, fontSizePt, bold, italic, strike,
                    underline, colorHex, vertAlign, inHyperlink, inField);
        }

        @Override public String toString() {
            return "StyleKey{font=" + fontFamily + ", size=" + fontSizePt +
                    ", b=" + bold + ", i=" + italic + ", s=" + strike +
                    ", u=" + underline + ", c=" + colorHex + ", va=" + vertAlign +
                    ", hyp=" + inHyperlink + ", fld=" + inField + '}';
        }
    }

    // ---------- 入口：文档分段 ----------
    public static List<Segment> segmentDocument(XWPFDocument doc) {
        List<Segment> out = new ArrayList<>();
        segmentBody(doc, out);
        for (XWPFHeader h : doc.getHeaderList()) segmentBody(h, out);
        for (XWPFFooter f : doc.getFooterList()) segmentBody(f, out);
        for (XWPFFootnote fn : doc.getFootnotes()) segmentBody(fn, out);
        return out;
    }

    /** 递归遍历 IBody（文档/单元格/页眉等） */
    private static void segmentBody(IBody body, List<Segment> out) {
        for (IBodyElement be : body.getBodyElements()) {
            if (be instanceof XWPFParagraph) {
                out.addAll(segmentParagraph((XWPFParagraph) be));

            } else if (be instanceof XWPFTable) {
                segmentTable((XWPFTable) be, out);

            } else if (be instanceof XWPFSDT) {
                // 兼容不同 POI 版本：优先反射 getBodyElements()；否则分别反射 getParagraphs()/getTables()
                ISDTContent content = ((XWPFSDT) be).getContent();
                boolean handled = false;
                try {
                    var m = content.getClass().getMethod("getBodyElements");
                    @SuppressWarnings("unchecked")
                    List<IBodyElement> bes = (List<IBodyElement>) m.invoke(content);
                    for (IBodyElement inner : bes) {
                        if (inner instanceof XWPFParagraph) out.addAll(segmentParagraph((XWPFParagraph) inner));
                        else if (inner instanceof XWPFTable) segmentTable((XWPFTable) inner, out);
                    }
                    handled = true;
                } catch (Throwable ignored) { /* fall through */ }

                if (!handled) {
                    try {
                        var mp = content.getClass().getMethod("getParagraphs");
                        @SuppressWarnings("unchecked")
                        List<XWPFParagraph> ps = (List<XWPFParagraph>) mp.invoke(content);
                        for (XWPFParagraph p : ps) out.addAll(segmentParagraph(p));
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

    /** 表格遍历（按行/单元格递归进入 cell 作为 IBody） */
    private static void segmentTable(XWPFTable t, List<Segment> out) {
        for (XWPFTableRow row : t.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                segmentBody(cell, out); // cell 实现 IBody
            }
        }
    }

    // ---------- 段落分段 ----------
    public static List<Segment> segmentParagraph(XWPFParagraph p) {
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
            if (i > 0 && (boundary || !k.equals(segKey))) {
                out.add(new Segment(p, segStart, i - 1, segKey, segText.toString()));
                segStart = i; segKey = k; segText.setLength(0);
            }
            appendRunTextPreserve(r, segText);
        }
        out.add(new Segment(p, segStart, runs.size() - 1, segKey, segText.toString()));
        return out;
    }

    // ---------- 指纹/边界/文本拼接 ----------
    public static StyleKey keyOf(XWPFRun r) {
        String family = r.getFontFamily();
        Integer sizePt = fontSizePtOf(r);              // 不用 getFontSize()，优先读 <w:sz>
        boolean bold = r.isBold();
        boolean italic = r.isItalic();
        boolean strike = r.isStrikeThrough();
        String underline = (r.getUnderline() != null) ? r.getUnderline().name() : null;
        String colorHex = r.getColor();
        String vertAlign = vertAlignOf(r);            // 用光标读 <w:vertAlign>
        boolean inHyperlink = (r instanceof XWPFHyperlinkRun);
        boolean inField = hasFieldMark(r.getCTR());
        return new StyleKey(family, sizePt, bold, italic, strike, underline, colorHex, vertAlign, inHyperlink, inField);
    }

    /** 读取字号（pt）：优先 <w:sz w:val> 的 half-points；取不到再安全回退 run.getFontSize() */
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
                            // 四舍五入转 pt
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
        } catch (Throwable ignore) {
            return null;
        }
    }

    /** 读取上下标：<w:vertAlign w:val="superscript|subscript|baseline"> */
    private static String vertAlignOf(XWPFRun r) {
        CTR ctr = r.getCTR();
        if (ctr == null) return null;
        try (XmlCursor c = ctr.newCursor()) {
            c.selectPath("declare namespace w='" + NS_W + "' .//w:vertAlign");
            if (c.toNextSelection()) {
                String val = c.getAttributeText(new QName(NS_W, "val"));
                return (val == null || val.isEmpty()) ? null : val;
            }
        }
        return null;
    }

    /** 运行是否在“域”里（用 sizeOfXXXArray 避免依赖具体 CT 类型） */
    private static boolean hasFieldMark(CTR ctr) {
        if (ctr == null) return false;
        if (ctr.sizeOfFldCharArray() > 0) return true;
        if (ctr.sizeOfInstrTextArray() > 0) return true;
        return false;
    }

    /** 硬边界：超链接容器变化 或 域状态变化 */
    private static boolean hardBoundary(XWPFRun cur, XWPFRun prev) {
        if ((cur instanceof XWPFHyperlinkRun) != (prev instanceof XWPFHyperlinkRun)) return true;
        boolean curFld = hasFieldMark(cur.getCTR());
        boolean prevFld = hasFieldMark(prev.getCTR());
        return curFld != prevFld;
    }

    /** 追加 run 文本（保留 <w:br/>→\\n、<w:tab/>→\\t） */
    private static void appendRunTextPreserve(XWPFRun r, StringBuilder sb) {
        CTR ctr = r.getCTR();
        if (ctr == null) return;

        for (CTText t : ctr.getTList()) {
            if (t == null) continue;
            String val = t.getStringValue();
            if (val != null) sb.append(val);
        }
        for (int i = 0; i < ctr.sizeOfBrArray(); i++) sb.append('\n');
        for (int i = 0; i < ctr.sizeOfTabArray(); i++) sb.append('\t');
    }
}
