package com.example.demo;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;

import javax.xml.namespace.QName;
import java.util.HashMap;
import java.util.Map;

/** 将合并后的文本写回到段内指定 run 范围，保留模板 run 样式并清理尾随 runs */
public final class SegmentRestorer {
    private SegmentRestorer() {}
    private static final String NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static final int MAX_T_CHUNK = 512; // 单个 <w:t> 最大字符数，超出则切片

    /** 模板选择策略 */
    public enum TemplateMode { FIRST_RUN, MAJORITY_STYLE, LONGEST_TEXT }

    /** 在段内按 [runFrom,runTo] 回填合并文本，使用 FIRST_RUN 模板 */
    public static void restoreSegmentInParagraph(XWPFParagraph p, int runFrom, int runTo, String mergedText) {
        restoreSegmentInParagraph(p, runFrom, runTo, mergedText, TemplateMode.FIRST_RUN);
    }

    /** 在段内按 [runFrom,runTo] 回填合并文本，模板模式可选 */
    public static void restoreSegmentInParagraph(XWPFParagraph p, int runFrom, int runTo, String mergedText, TemplateMode mode) {
        if (p == null || runFrom < 0 || runTo >= p.getRuns().size() || runFrom > runTo) return;

        int baseIdx = pickTemplateIndex(p, runFrom, runTo, mode);
        if (baseIdx != runFrom) copyRPr(p.getRuns().get(baseIdx), p.getRuns().get(runFrom));

        writeTextToRun(p.getRuns().get(runFrom), mergedText);

        for (int i = runTo; i >= runFrom + 1; i--) {
            XWPFRun r = p.getRuns().get(i);
            if (isAnchoredRun(r)) clearRunText(r); else p.removeRun(i);
        }
    }

    /** 清理段内所有 proofErr（语法/拼写范围标记） */
    public static void stripProofErr(XWPFParagraph p) {
        if (p == null) return;
        try (XmlCursor c = p.getCTP().newCursor()) {
            c.selectPath("declare namespace w='" + NS_W + "' .//w:proofErr");
            while (c.toNextSelection()) c.removeXml();
        } catch (Exception ignore) {}
    }

    // —— 工具 —— //

    private static void copyRPr(XWPFRun src, XWPFRun dst) {
        var s = src.getCTR(); var d = dst.getCTR();
        if (s == null || d == null) return;
        if (!d.isSetRPr()) d.addNewRPr();
        if (s.isSetRPr()) d.getRPr().set(s.getRPr());
    }

    private static boolean isAnchoredRun(XWPFRun r) {
        var ctr = r.getCTR(); if (ctr == null) return false;
        if (r instanceof XWPFHyperlinkRun) return true;
        if (ctr.sizeOfFldCharArray() > 0) return true;
        if (ctr.sizeOfInstrTextArray() > 0) return true;
        if (ctr.sizeOfFootnoteReferenceArray() > 0) return true;
        if (ctr.sizeOfCommentReferenceArray() > 0) return true;
        return false;
    }

    private static void clearRunText(XWPFRun r) {
        var ctr = r.getCTR(); if (ctr == null) return;
        try (XmlCursor c = ctr.newCursor()) {
            c.selectPath("declare namespace w='" + NS_W + "' ./w:t|./w:br|./w:cr|./w:tab|./w:instrText");
            while (c.toNextSelection()) c.removeXml();
        }
    }

    private static void writeTextToRun(XWPFRun r, String text) {
        if (text == null) text = "";
        String s = text.replace("\r\n","\n").replace('\r','\n')
                       .replace('\u2028','\n').replace('\u2029','\n');

        clearRunText(r);
        var ctr = r.getCTR();
        String[] lines = s.split("\n", -1);
        try (XmlCursor c = ctr.newCursor()) {
            c.toEndToken();
            for (int i = 0; i < lines.length; i++) {
                if (i > 0) { c.beginElement(new QName(NS_W, "br")); c.toParent(); }
                String line = lines[i] == null ? "" : lines[i];
                int off = 0, n = line.length();
                if (n == 0) {
                    c.beginElement(new QName(NS_W, "t"));
                    c.insertAttributeWithValue(new QName("http://www.w3.org/XML/1998/namespace","space","xml"), "preserve");
                    c.insertChars("");
                    c.toParent();
                } else {
                    while (off < n) {
                        int end = Math.min(off + MAX_T_CHUNK, n);
                        c.beginElement(new QName(NS_W, "t"));
                        c.insertAttributeWithValue(new QName("http://www.w3.org/XML/1998/namespace","space","xml"), "preserve");
                        c.insertChars(line.substring(off, end));
                        c.toParent();
                        off = end;
                    }
                }
            }
        }
    }

    private static int pickTemplateIndex(XWPFParagraph p, int from, int to, TemplateMode mode) {
        if (mode == TemplateMode.FIRST_RUN) return from;
        if (mode == TemplateMode.LONGEST_TEXT) {
            int best = from, bestLen = -1;
            for (int i = from; i <= to; i++) {
                int len = runTextLength(p.getRuns().get(i));
                if (len > bestLen) { bestLen = len; best = i; }
            }
            return best;
        }
        Map<String,Integer> freq = new HashMap<>();
        int best = from, bestScore = -1;
        for (int i = from; i <= to; i++) {
            XWPFRun r = p.getRuns().get(i);
            int size = -1; try { @SuppressWarnings("deprecation") int fs=r.getFontSize(); size=(fs>0?fs:-1);} catch(Throwable ignore){}
            String key = (r.isBold()?"B":"b")+(r.isItalic()?"I":"i")+(r.getUnderline()!=null?r.getUnderline().name():"u-")+":"+size;
            int sc = freq.merge(key,1,Integer::sum);
            if (sc > bestScore) { bestScore = sc; best = i; }
        }
        return best;
    }

    private static int runTextLength(XWPFRun r) {
        var ctr = r.getCTR(); if (ctr == null) return 0;
        int n = 0;
        for (var t : ctr.getTList()) { if (t!=null && t.getStringValue()!=null) n += t.getStringValue().length(); }
        n += ctr.sizeOfBrArray();
        n += ctr.sizeOfTabArray();
        return n;
    }
}
