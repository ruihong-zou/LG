package com.example.demo;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

import javax.xml.namespace.QName;
import java.util.Arrays;

/** 原位写回工具：把一段 runs 合并写回到“首 run”，保留首 run 样式，安全清理后续 runs */
public final class SegmentRestorer {

    private SegmentRestorer() {}

    private static final QName QN_XML_SPACE = new QName("http://www.w3.org/XML/1998/namespace", "space", "xml");

    /**
     * 将 [runFrom..runTo] 覆盖为 mergedText：
     * - 文本全部写入第一个 run（runFrom），保留其 rPr（即“保留首 run 的格式”）；
     * - 中间与尾部 run 若无锚点可删除；有锚点则仅清空文字；
     * - 支持 \n -> <w:br/>, \t -> <w:tab/>，并为 <w:t> 设置 xml:space="preserve"。
     */
    public static void restoreSegmentInParagraph(XWPFParagraph p, int runFrom, int runTo, String mergedText) {
        if (p == null) return;
        if (runFrom < 0) runFrom = 0;
        if (runTo >= p.getRuns().size()) runTo = p.getRuns().size() - 1;
        if (runFrom > runTo) return;

        XWPFRun base = p.getRuns().get(runFrom);
        writeTextPreserve(base, mergedText);

        // 从后向前处理其他 run，避免索引错位，并规避 XmlValueDisconnectedException
        for (int i = runTo; i > runFrom; i--) {
            XWPFRun r = p.getRuns().get(i);
            if (hasAnchors(r)) {
                clearRunText(r);
            } else {
                try {
                    p.removeRun(i);
                } catch (Throwable ignore) {
                    clearRunText(r);
                }
            }
        }
    }

    /** 同上，但用于单元格里的段落（便捷封装） */
    public static void restoreSegmentInCell(XWPFTableCell cell, int paraIdx, int runFrom, int runTo, String mergedText) {
        if (cell == null) return;
        if (paraIdx < 0 || paraIdx >= cell.getParagraphs().size()) return;
        restoreSegmentInParagraph(cell.getParagraphs().get(paraIdx), runFrom, runTo, mergedText);
    }

    // =============== 低层工具 ===============

    /** 将纯文本写入 run：清空现有 t/br/tab/instrText，按 \n 生成 <w:br/>，并设置 xml:space="preserve"。 */
    public static void writeTextPreserve(XWPFRun run, String text) {
        if (run == null) return;
        String s = (text == null) ? "" : text;
        s = s.replace("\r\n", "\n").replace('\r', '\n').replace('\u2028', '\n').replace('\u2029', '\n');

        clearRunText(run);

        String[] lines = s.split("\n", -1);
        lines = trimTrailingEmpty(lines);

        CTR ctr = run.getCTR();
        for (int i = 0; i < lines.length; i++) {
            if (i > 0) ctr.addNewBr();
            CTText t = ctr.addNewT();
            t.setStringValue(lines[i] == null ? "" : lines[i]);
            try (org.apache.xmlbeans.XmlCursor tc = t.newCursor()) { tc.setAttributeText(QN_XML_SPACE, "preserve"); } catch (Throwable ignore) {}
        }
    }

    /** 清空一个 run 下的 t/br/cr/tab/instrText 文本节点，不动 rPr 等属性节点。 */
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

    /**
     * 是否含“结构锚点”：超链接容器、域边界/指令、脚注/批注引用等；
     * 含锚点则不删除 run，只清空文字。
     */
    public static boolean hasAnchors(XWPFRun run) {
        if (run == null) return false;
        if (run instanceof XWPFHyperlinkRun) return true;
        CTR ctr = run.getCTR(); if (ctr == null) return false;
        if (!ctr.getFldCharList().isEmpty()) return true;     // 域 begin/separate/end
        if (!ctr.getInstrTextList().isEmpty()) return true;   // 域指令
        if (!ctr.getFootnoteReferenceList().isEmpty()) return true;
        if (!ctr.getEndnoteReferenceList().isEmpty()) return true;
        if (!ctr.getCommentReferenceList().isEmpty()) return true;
        return false;
    }

    private static String[] trimTrailingEmpty(String[] arr) {
        int end = arr.length;
        while (end > 1 && (arr[end - 1] == null || arr[end - 1].isEmpty())) end--;
        return Arrays.copyOf(arr, end);
    }
}
