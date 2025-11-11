package com.example.demo;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

import javax.xml.namespace.QName;
import java.util.Arrays;

/** 段级原位回填：把 runFrom..runTo 的文本合并写入首 run，保留样式并处理换行/制表/空格 */
public final class SegmentRestorer {

    private SegmentRestorer() {}

    private static final QName QN_XML_SPACE = new QName("http://www.w3.org/XML/1998/namespace", "space", "xml");

    /** 将 [runFrom..runTo] 覆盖为 mergedText，写入首 run 并清理其余 run */
    public static void restoreSegmentInParagraph(XWPFParagraph p, int runFrom, int runTo, String mergedText) {
        if (p == null) return;
        if (runFrom < 0) runFrom = 0;
        if (runTo >= p.getRuns().size()) runTo = p.getRuns().size() - 1;
        if (runFrom > runTo) return;

        XWPFRun base = p.getRuns().get(runFrom);
        writeTextPreserve(base, mergedText);

        for (int i = runTo; i > runFrom; i--) {
            XWPFRun r = p.getRuns().get(i);
            if (hasAnchors(r)) {
                clearRunText(r);
            } else {
                try { p.removeRun(i); } catch (Throwable ignore) { clearRunText(r); }
            }
        }
    }

    /** 将单元格指定段落的 [runFrom..runTo] 覆盖为 mergedText */
    public static void restoreSegmentInCell(XWPFTableCell cell, int paraIdx, int runFrom, int runTo, String mergedText) {
        if (cell == null) return;
        if (paraIdx < 0 || paraIdx >= cell.getParagraphs().size()) return;
        restoreSegmentInParagraph(cell.getParagraphs().get(paraIdx), runFrom, runTo, mergedText);
    }

    /** 将文本写入 run，处理换行/制表并设置 xml:space="preserve" */
    public static void writeTextPreserve(XWPFRun run, String text) {
        if (run == null) return;
        String s = (text == null) ? "" : text;
        s = s.replace("\r\n", "\n").replace('\r', '\n')
             .replace('\u2028', '\n').replace('\u2029', '\n');

        clearRunText(run);

        String[] lines = s.split("\n", -1);
        lines = trimTrailingEmpty(lines);

        CTR ctr = run.getCTR();
        for (int i = 0; i < lines.length; i++) {
            if (i > 0) ctr.addNewBr();
            appendLineWithTabs(ctr, lines[i]);
        }
    }

    /** 清空 run 的文本子节点（保留 rPr 等属性） */
    public static void clearRunText(XWPFRun run) {
        if (run == null) return;
        CTR ctr = run.getCTR(); if (ctr == null) return;
        try (org.apache.xmlbeans.XmlCursor c = ctr.newCursor()) {
            if (c.toFirstChild()) {
                do {
                    javax.xml.namespace.QName n = c.getName(); if (n == null) continue;
                    String ln = n.getLocalPart();
                    if ("t".equals(ln) || "br".equals(ln) || "cr".equals(ln) || "tab".equals(ln) || "instrText".equals(ln)) {
                        c.removeXml();
                    }
                } while (c.toNextSibling());
            }
        }
    }

    /** 判断 run 是否含结构性锚点（域、超链接、脚注/批注引用） */
    public static boolean hasAnchors(XWPFRun run) {
        if (run == null) return false;
        if (run instanceof XWPFHyperlinkRun) return true;
        CTR ctr = run.getCTR(); if (ctr == null) return false;
        if (!ctr.getFldCharList().isEmpty()) return true;
        if (!ctr.getInstrTextList().isEmpty()) return true;
        if (!ctr.getFootnoteReferenceList().isEmpty()) return true;
        if (!ctr.getEndnoteReferenceList().isEmpty()) return true;
        if (!ctr.getCommentReferenceList().isEmpty()) return true;
        return false;
    }

    /** 将一行中以 \t 分隔的文本写为多个 <w:t xml:space="preserve">，并在段间插入 <w:tab/> */
    private static void appendLineWithTabs(CTR ctr, String line) {
        String[] parts = line.split("\t", -1);
        for (int j = 0; j < parts.length; j++) {
            CTText t = ctr.addNewT();
            t.setStringValue(parts[j] == null ? "" : parts[j]);
            try (org.apache.xmlbeans.XmlCursor tc = t.newCursor()) {
                tc.setAttributeText(QN_XML_SPACE, "preserve");
            } catch (Throwable ignore) {}
            if (j < parts.length - 1) ctr.addNewTab();
        }
    }

    private static String[] trimTrailingEmpty(String[] arr) {
        int end = arr.length;
        while (end > 1 && (arr[end - 1] == null || arr[end - 1].isEmpty())) end--;
        return Arrays.copyOf(arr, end);
    }
}
