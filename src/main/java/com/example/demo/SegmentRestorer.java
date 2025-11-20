// File: src/main/java/com/example/demo/SegmentRestorer.java
package com.example.demo;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

import java.util.HashMap;
import java.util.Map;

import javax.xml.namespace.QName;

/** 段内按 [runFrom, runTo] 回填合并文本；含 <w:tab/> 的 run 作为锚点保留且不写入、不删除 */
public final class SegmentRestorer {
    private SegmentRestorer() {}
    private static final String NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public enum TemplateMode { FIRST_RUN, MAJORITY_STYLE, LONGEST_TEXT }

    public static void restoreSegmentInParagraph(XWPFParagraph p, int runFrom, int runTo, String mergedText) {
        restoreSegmentInParagraph(p, runFrom, runTo, mergedText, TemplateMode.FIRST_RUN);
    }

    public static void restoreSegmentInParagraph(XWPFParagraph p, int runFrom, int runTo, String mergedText, TemplateMode mode) {
        if (p == null || runFrom < 0 || runTo >= p.getRuns().size() || runFrom > runTo) return;

        int baseIdx = pickTemplateIndexNonTab(p, runFrom, runTo, mode);
        if (baseIdx < runFrom || baseIdx > runTo) return; // 没有可写入的非 tab run

        writeTextToRun(p.getRuns().get(baseIdx), mergedText);

        for (int i = runTo; i >= runFrom; i--) {
            if (i == baseIdx) continue;
            XWPFRun r = p.getRuns().get(i);
            if (isAnchoredRun(r)) clearRunText(r); // 锚点：保留结构，只清空文本
            else p.removeRun(i);
        }
    }

    public static void stripProofErr(XWPFParagraph p) {
        if (p == null) return;
        try (XmlCursor c = p.getCTP().newCursor()) {
            c.selectPath("declare namespace w='" + NS_W + "' .//w:proofErr");
            while (c.toNextSelection()) c.removeXml();
        } catch (Exception ignore) {}
    }

    private static void copyRPr(XWPFRun src, XWPFRun dst) {
        var s = src.getCTR(); var d = dst.getCTR();
        if (s == null || d == null) return;
        if (!d.isSetRPr()) d.addNewRPr();
        if (s.isSetRPr()) d.getRPr().set(s.getRPr());
    }

    /** 含超链/域/脚注/批注 或 含 <w:tab/> 的 run 视为锚点 */
    private static boolean isAnchoredRun(XWPFRun r) {
        var ctr = r.getCTR(); if (ctr == null) return false;
        if (r instanceof XWPFHyperlinkRun) return true;
        if (ctr.sizeOfFldCharArray() > 0) return true;
        if (ctr.sizeOfInstrTextArray() > 0) return true;
        if (ctr.sizeOfFootnoteReferenceArray() > 0) return true;
        if (ctr.sizeOfCommentReferenceArray() > 0) return true;
        if (ctr.sizeOfTabArray() > 0) return true; // tab 锚点
        return false;
    }

    private static void clearRunText(XWPFRun r) {
        var ctr = r.getCTR(); if (ctr == null) return;
        try (XmlCursor c = ctr.newCursor()) {
            c.selectPath("declare namespace w='" + NS_W + "' ./w:t|./w:br|./w:cr|./w:instrText");
            while (c.toNextSelection()) c.removeXml();
        }
    }

    /** 每一行写入一个 <w:t xml:space="preserve">；不处理 \t，避免触碰 tab run */
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
                c.beginElement(new QName(NS_W, "t"));
                c.insertAttributeWithValue(new QName("http://www.w3.org/XML/1998/namespace","space","xml"), "preserve");
                c.insertChars(lines[i] == null ? "" : lines[i]);
                c.toParent();
            }
        }
    }

    private static int pickTemplateIndexNonTab(XWPFParagraph p, int from, int to, TemplateMode mode) {
        // 先尝试策略选择，跳过含 tab 的 run
        if (mode == TemplateMode.LONGEST_TEXT) {
            int best = -1, bestLen = -1;
            for (int i = from; i <= to; i++) {
                XWPFRun r = p.getRuns().get(i);
                if (hasTab(r)) continue;
                int len = runTextLength(r);
                if (len > bestLen) { bestLen = len; best = i; }
            }
            if (best != -1) return best;
        } else if (mode == TemplateMode.MAJORITY_STYLE) {
            Map<String,Integer> freq = new HashMap<>();
            int best = -1, bestScore = -1;
            for (int i = from; i <= to; i++) {
                XWPFRun r = p.getRuns().get(i);
                if (hasTab(r)) continue;
                int size = -1; try { @SuppressWarnings("deprecation") int fs=r.getFontSize(); size=(fs>0?fs:-1);} catch(Throwable ignore){}
                String key = (r.isBold()?"B":"b")+(r.isItalic()?"I":"i")+(r.getUnderline()!=null?r.getUnderline().name():"u-")+":"+size;
                int sc = freq.merge(key,1,Integer::sum);
                if (sc > bestScore) { bestScore = sc; best = i; }
            }
            if (best != -1) return best;
        }
        // 默认取区间内第一个“非 tab” run
        for (int i = from; i <= to; i++) if (!hasTab(p.getRuns().get(i))) return i;
        return -1;
    }

    private static boolean hasTab(XWPFRun r) {
        CTR ctr = r.getCTR(); return ctr != null && ctr.sizeOfTabArray() > 0;
    }

    private static int runTextLength(XWPFRun r) {
        var ctr = r.getCTR(); if (ctr == null) return 0;
        int n = 0;
        for (var t : ctr.getTList()) { if (t!=null && t.getStringValue()!=null) n += t.getStringValue().length(); }
        n += ctr.sizeOfBrArray();
        return n;
    }
}
