package com.example.demo;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.apache.xmlbeans.XmlCursor;

import javax.xml.namespace.QName;
import java.util.*;

/** 按“格式变化”对段落中的 runs 进行分段，输出段级文本与 run 范围 */
public final class FormatChangeSegmenter {

    private FormatChangeSegmenter() {}

    private static final String NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /** 代表一个段内的格式一致片段 */
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

    /** 片段的样式指纹 */
    public static final class StyleKey {
        public final String fontFamily;
        public final Integer fontSizePt;
        public final boolean bold;
        public final boolean italic;
        public final boolean strike;
        public final String underline;
        public final String colorHex;
        public final String vertAlign;
        public final boolean inHyperlink;
        public final boolean inField;

        public StyleKey(String fontFamily, Integer fontSizePt, boolean bold, boolean italic, boolean strike,
                        String underline, String colorHex, String vertAlign, boolean inHyperlink, boolean inField) {
            this.fontFamily = nz(fontFamily);
            this.fontSizePt = (fontSizePt != null && fontSizePt >= 0) ? fontSizePt : null;
            this.bold = bold; this.italic = italic; this.strike = strike;
            this.underline = nz(underline); this.colorHex = normHex(colorHex);
            this.vertAlign = nz(vertAlign); this.inHyperlink = inHyperlink; this.inField = inField;
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
            if (this == o) return true; if (!(o instanceof StyleKey)) return false; StyleKey k = (StyleKey)o;
            return bold==k.bold && italic==k.italic && strike==k.strike && inHyperlink==k.inHyperlink && inField==k.inField &&
                    Objects.equals(fontFamily,k.fontFamily) && Objects.equals(fontSizePt,k.fontSizePt) &&
                    Objects.equals(underline,k.underline) && Objects.equals(colorHex,k.colorHex) &&
                    Objects.equals(vertAlign,k.vertAlign);
        }
        @Override public int hashCode() {
            return Objects.hash(fontFamily,fontSizePt,bold,italic,strike,underline,colorHex,vertAlign,inHyperlink,inField);
        }
    }

    /** 对文档进行分段（正文、页眉页脚、脚注） */
    public static List<Segment> segmentDocument(XWPFDocument doc) {
        List<Segment> out = new ArrayList<>();

        // 正文
        segmentBody(doc, out);

        // 页眉/页脚
        for (XWPFHeader h : doc.getHeaderList()) {
            segmentBody(h, out);
        }
        for (XWPFFooter f : doc.getFooterList()) {
            segmentBody(f, out);
        }

        // 脚注
        for (XWPFFootnote fn : doc.getFootnotes()) {
            segmentBody(fn, out);
        }

        // 端注（有的版本存在，有的没有；用反射兼容）
        try {
            var m = XWPFDocument.class.getMethod("getEndnotes");
            @SuppressWarnings("unchecked")
            List<Object> endnotes = (List<Object>) m.invoke(doc);
            for (Object en : endnotes) {
                if (en instanceof IBody) segmentBody((IBody) en, out);
            }
        } catch (Throwable ignored) {
            // 没有端注 API 就跳过
        }

        return out;
    }

    /** 遍历 IBody 内容（段落/表格/SDT） */
    private static void segmentBody(IBody body, List<Segment> out) {
        for (IBodyElement be : body.getBodyElements()) {
            if (be instanceof XWPFParagraph) {
                out.addAll(segmentParagraph((XWPFParagraph) be));
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
                        if (inner instanceof XWPFParagraph) out.addAll(segmentParagraph((XWPFParagraph) inner));
                        else if (inner instanceof XWPFTable) segmentTable((XWPFTable) inner, out);
                    }
                    handled = true;
                } catch (Throwable ignored) {}
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

    /** 遍历表格并递归进入单元格 */
    private static void segmentTable(XWPFTable t, List<Segment> out) {
        for (XWPFTableRow row : t.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                segmentBody(cell, out);
            }
        }
    }

    /** 对单个段落按“格式变化”分段 */
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

    /** 生成样式指纹 */
    public static StyleKey keyOf(XWPFRun r) {
        String family = r.getFontFamily();
        Integer sizePt = fontSizePtOf(r);
        boolean bold = r.isBold();
        boolean italic = r.isItalic();
        boolean strike = r.isStrikeThrough();
        String underline = (r.getUnderline() != null) ? r.getUnderline().name() : null;
        String colorHex = r.getColor();
        String vertAlign = vertAlignOf(r);
        boolean inHyperlink = (r instanceof XWPFHyperlinkRun);
        boolean inField = hasFieldMark(r.getCTR());
        return new StyleKey(family, sizePt, bold, italic, strike, underline, colorHex, vertAlign, inHyperlink, inField);
    }

    /** 读取字号（pt）：优先解析 <w:sz w:val>；取不到再回退 run.getFontSize() */
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

    /** 读取上下标：解析 <w:vertAlign w:val="superscript|subscript|baseline"> */
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

    /** 判断 run 是否处于域中 */
    private static boolean hasFieldMark(CTR ctr) {
        if (ctr == null) return false;
        if (ctr.sizeOfFldCharArray() > 0) return true;
        if (ctr.sizeOfInstrTextArray() > 0) return true;
        return false;
    }

    /** 硬边界：超链接容器变化或域状态变化 */
    private static boolean hardBoundary(XWPFRun cur, XWPFRun prev) {
        if ((cur instanceof XWPFHyperlinkRun) != (prev instanceof XWPFHyperlinkRun)) return true;
        boolean curFld = hasFieldMark(cur.getCTR());
               boolean prevFld = hasFieldMark(prev.getCTR());
        return curFld != prevFld;
    }

    /** 追加 run 文本到 StringBuilder，严格按子节点顺序保留换行与制表 */
    private static void appendRunTextPreserve(XWPFRun r, StringBuilder sb) {
        CTR ctr = r.getCTR();
        if (ctr == null) return;

        try (org.apache.xmlbeans.XmlCursor c = ctr.newCursor()) {
            if (!c.toFirstChild()) return;
            do {
                javax.xml.namespace.QName n = c.getName();
                if (n == null) continue;
                String ln = n.getLocalPart();

                switch (ln) {
                    case "t":
                    case "instrText": {
                        String v = c.getTextValue();
                        if (v != null) sb.append(v);
                        break;
                    }
                    case "br":
                    case "cr": {
                        sb.append('\n');
                        break;
                    }
                    case "tab": {
                        sb.append('\t');
                        break;
                    }
                    case "softHyphen": {
                        sb.append('\u00AD'); // 可选软连字符
                        break;
                    }
                    case "noBreakHyphen": {
                        sb.append('\u2011'); // 不换行连字符
                        break;
                    }
                    default:
                        // 其他子节点忽略
                        break;
                }
            } while (c.toNextSibling());
        }
    }

}
