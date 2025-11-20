// File: src/main/java/com/example/demo/WordDocxExtractorRestorer.java
package com.example.demo;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;

import javax.xml.namespace.QName;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.*;
import java.util.regex.Pattern;

/** 支持：段片提取/回填；文本框 run 提取/回填；外层 SDT 回填。含 <w:tab/> 的 run 作为硬边界保留。 */
public class WordDocxExtractorRestorer {

    public static class TextElement {
        public final String text;
        public final String type; // "paraSeg" | "cellSeg" | "docxTextBoxRun" | "docxSdtField"
        public final Map<String, Object> position;

        public TextElement(String text, String type, Map<String, Object> position) {
            this.text = text; this.type = type; this.position = position;
        }
    }

    private static class XmlChange {
        final String type;   // "docxTextBoxRun" | "docxSdtField"
        final String part;
        final String sdtPath;
        final String boxPath;
        final String alias;
        final String tag;
        final String kind;
        final String newText;
        final Integer runOrd;

        XmlChange(String type, String part, String sdtPath, String boxPath,
                  String alias, String tag, String kind, String newText, Integer runOrd) {
            this.type = type; this.part = part; this.sdtPath = sdtPath; this.boxPath = boxPath;
            this.alias = alias; this.tag = tag; this.kind = kind; this.newText = newText; this.runOrd = runOrd;
        }
    }

    private static final String NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static final QName QN_W_SDT          = new QName(NS_W, "sdt");
    private static final QName QN_W_SDTCONTENT   = new QName(NS_W, "sdtContent");
    private static final QName QN_W_SDTPr        = new QName(NS_W, "sdtPr");
    private static final QName QN_W_ALIAS        = new QName(NS_W, "alias");
    private static final QName QN_W_TAG          = new QName(NS_W, "tag");
    private static final QName QN_W_VAL          = new QName(NS_W, "val");
    private static final QName QN_W_TXBX_CONTENT = new QName(NS_W, "txbxContent");
    private static final QName QN_XML_SPACE      = new QName("http://www.w3.org/XML/1998/namespace", "space", "xml");

    private static final String PARA_SEP = "\u2029";
    private static final Pattern ANY_BREAK = Pattern.compile("\r\n|\r|\n|\u2028|\u2029|\u000B|\u000C|\u0085");

    public static List<TextElement> extractWordTexts(XWPFDocument doc) {
        List<TextElement> elements = new ArrayList<>();

        // 正文段片
        List<XWPFParagraph> paras = doc.getParagraphs();
        for (int pIdx = 0; pIdx < paras.size(); pIdx++) {
            XWPFParagraph p = paras.get(pIdx);
            List<FormatChangeSegmenter.Segment> segs = FormatChangeSegmenter.segmentParagraph(p, MergePolicy.loose());
            for (FormatChangeSegmenter.Segment seg : segs) {
                if (notBlank(seg.text)) {
                    Map<String, Object> pos = new HashMap<>();
                    pos.put("paragraphIndex", pIdx);
                    pos.put("runStartIndex", seg.runStartIdx);
                    pos.put("runEndIndex", seg.runEndIdx);
                    elements.add(new TextElement(seg.text, "paraSeg", pos));
                }
            }
        }

        // 表格段片
        for (int tIdx = 0; tIdx < doc.getTables().size(); tIdx++) {
            XWPFTable t = doc.getTables().get(tIdx);
            for (int r = 0; r < t.getRows().size(); r++) {
                XWPFTableRow row = t.getRows().get(r);
                for (int c = 0; c < row.getTableCells().size(); c++) {
                    XWPFTableCell cell = row.getTableCells().get(c);
                    List<XWPFParagraph> ps = cell.getParagraphs();
                    for (int pi = 0; pi < ps.size(); pi++) {
                        XWPFParagraph p = ps.get(pi);
                        List<FormatChangeSegmenter.Segment> segs = FormatChangeSegmenter.segmentParagraph(p, MergePolicy.loose());
                        for (FormatChangeSegmenter.Segment seg : segs) {
                            if (notBlank(seg.text)) {
                                Map<String, Object> pos = new HashMap<>();
                                pos.put("tableIndex", tIdx);
                                pos.put("rowIndex", r);
                                pos.put("cellIndex", c);
                                pos.put("paraInCell", pi);
                                pos.put("runStartIndex", seg.runStartIdx);
                                pos.put("runEndIndex", seg.runEndIdx);
                                elements.add(new TextElement(seg.text, "cellSeg", pos));
                            }
                        }
                    }
                }
            }
        }

        // 文本框 run（跳过含 tab 的 run） + 文本框外 SDT
        collectDocXmlContainers(doc.getDocument(), elements);
        return elements;
    }

    public static void restoreWordTexts(XWPFDocument doc, List<TextElement> elements, List<String> translated) {
        Map<Integer, List<Integer>> paraGroups = new HashMap<>();
        Map<String, List<Integer>> cellGroups = new HashMap<>();

        for (int i = 0; i < elements.size(); i++) {
            TextElement el = elements.get(i);
            if ("paraSeg".equals(el.type)) {
                int p = (Integer) el.position.get("paragraphIndex");
                paraGroups.computeIfAbsent(p, k -> new ArrayList<>()).add(i);
            } else if ("cellSeg".equals(el.type)) {
                String key = el.position.get("tableIndex") + "|" + el.position.get("rowIndex") + "|" +
                             el.position.get("cellIndex") + "|" + el.position.get("paraInCell");
                cellGroups.computeIfAbsent(key, k -> new ArrayList<>()).add(i);
            }
        }

        for (Map.Entry<Integer, List<Integer>> e : paraGroups.entrySet()) {
            int pIdx = e.getKey();
            XWPFParagraph p = doc.getParagraphs().get(pIdx);

            e.getValue().sort((i1, i2) -> {
                int s1 = (Integer) elements.get(i1).position.get("runStartIndex");
                int s2 = (Integer) elements.get(i2).position.get("runStartIndex");
                return Integer.compare(s2, s1);
            });

            for (int idx : e.getValue()) {
                TextElement el = elements.get(idx);
                int s = (Integer) el.position.get("runStartIndex");
                int rEnd = (Integer) el.position.get("runEndIndex");
                SegmentRestorer.restoreSegmentInParagraph(p, s, rEnd, translated.get(idx));
            }
            SegmentRestorer.stripProofErr(p);
        }

        for (Map.Entry<String, List<Integer>> e : cellGroups.entrySet()) {
            String[] parts = e.getKey().split("\\|");
            int ti = Integer.parseInt(parts[0]);
            int ri = Integer.parseInt(parts[1]);
            int ci = Integer.parseInt(parts[2]);
            int pi = Integer.parseInt(parts[3]);

            XWPFParagraph p = doc.getTables().get(ti).getRow(ri).getCell(ci).getParagraphArray(pi);

            e.getValue().sort((i1, i2) -> {
                int s1 = (Integer) elements.get(i1).position.get("runStartIndex");
                int s2 = (Integer) elements.get(i2).position.get("runStartIndex");
                return Integer.compare(s2, s1);
            });

            for (int idx : e.getValue()) {
                TextElement el = elements.get(idx);
                int s = (Integer) el.position.get("runStartIndex");
                int rEnd = (Integer) el.position.get("runEndIndex");
                SegmentRestorer.restoreSegmentInParagraph(p, s, rEnd, translated.get(idx));
            }
            SegmentRestorer.stripProofErr(p);
        }

        Map<String, List<XmlChange>> changesByPart = new HashMap<>();
        for (int i = 0; i < elements.size(); i++) {
            TextElement el = elements.get(i);
            String typ = el.type == null ? "" : el.type;
            if ("docxTextBoxRun".equalsIgnoreCase(typ)) {
                String part    = asString(el.position.get("part"), "/word/document.xml");
                String sdtPath = asString(el.position.get("sdtPath"), "-");
                String boxPath = asString(el.position.get("boxPath"), "-");
                Integer runOrd = (Integer) el.position.get("runOrd");
                XmlChange ch = new XmlChange("docxTextBoxRun", part, sdtPath, boxPath,
                        "", "", "w:txbxContent", translated.get(i), runOrd);
                changesByPart.computeIfAbsent(part, k -> new ArrayList<>()).add(ch);
            } else if ("docxSdtField".equalsIgnoreCase(typ)) {
                String part    = asString(el.position.get("part"), "/word/document.xml");
                String sdtPath = asString(el.position.get("sdtPath"), "-");
                String boxPath = asString(el.position.get("boxPath"), "-");
                String alias   = asString(el.position.get("alias"), "");
                String tag     = asString(el.position.get("tag"), "");
                XmlChange ch = new XmlChange("docxSdtField", part, sdtPath, boxPath,
                        alias, tag, "sdt", translated.get(i), null);
                changesByPart.computeIfAbsent(part, k -> new ArrayList<>()).add(ch);
            }
        }

        List<XmlChange> docChanges = changesByPart.get("/word/document.xml");
        if (docChanges != null && !docChanges.isEmpty()) {
            applyChangesToPart(doc.getDocument(), docChanges);
        }
    }

    private static void collectDocXmlContainers(XmlObject root, List<TextElement> out) {
        Deque<Integer> sdtStack = new ArrayDeque<>();
        Deque<Integer> boxStack = new ArrayDeque<>();
        Map<Integer, Integer> boxDepthCounters = new HashMap<>();
        int[] sdtCounter = new int[]{0};
        dfsCollect(root, sdtStack, boxStack, boxDepthCounters, sdtCounter, out);
    }

    private static void dfsCollect(XmlObject node, Deque<Integer> sdtStack, Deque<Integer> boxStack,
                                   Map<Integer,Integer> boxDepthCounters, int[] sdtCounter, List<TextElement> out) {
        try (XmlCursor cur = node.newCursor()) {
            if (!cur.toFirstChild()) return;
            do {
                QName name = cur.getName();
                XmlObject child = cur.getObject();

                if (name != null) {
                    if (QN_W_SDT.equals(name)) {
                        int idx = ++sdtCounter[0];
                        sdtStack.push(idx);

                        boolean insideTextBox = hasAncestor(child, QN_W_TXBX_CONTENT);
                        if (!insideTextBox) {
                            SdtMeta meta = readSdtMeta(child);
                            String text = aggregateParagraphSeparated(child, QN_W_SDTCONTENT);
                            if (notBlank(text)) {
                                Map<String, Object> pos = new HashMap<>();
                                pos.put("type", "docxSdtField");
                                pos.put("part", "/word/document.xml");
                                pos.put("sdtPath", pathString(sdtStack));
                                pos.put("boxPath", pathString(boxStack));
                                pos.put("alias", meta.alias);
                                pos.put("tag", meta.tag);
                                pos.put("kind", "sdt");
                                out.add(new TextElement(text, "docxSdtField", pos));
                            }
                        }

                        dfsCollect(child, sdtStack, boxStack, boxDepthCounters, sdtCounter, out);
                        sdtStack.pop();
                        continue;
                    }

                    if (QN_W_TXBX_CONTENT.equals(name)) {
                        int depth = boxStack.size();
                        int next = boxDepthCounters.getOrDefault(depth, 0) + 1;
                        boxDepthCounters.put(depth, next);
                        boxStack.push(next);

                        try (XmlCursor rc = child.newCursor()) {
                            rc.selectPath("declare namespace w='" + NS_W + "' .//w:r");
                            int runOrd = 0;
                            while (rc.toNextSelection()) {
                                XmlObject rObj = rc.getObject();
                                if (runHasTabXml(rObj)) { runOrd++; continue; } // 跳过含 tab 的 run
                                String txt = getRunTextPreserveBrTab(rObj);
                                if (notBlank(txt)) {
                                    Map<String, Object> pos = new HashMap<>();
                                    pos.put("type", "docxTextBoxRun");
                                    pos.put("part", "/word/document.xml");
                                    pos.put("sdtPath", pathString(sdtStack));
                                    pos.put("boxPath", pathString(boxStack));
                                    pos.put("runOrd", runOrd);
                                    out.add(new TextElement(txt, "docxTextBoxRun", pos));
                                }
                                runOrd++;
                            }
                        }

                        dfsCollect(child, sdtStack, boxStack, boxDepthCounters, sdtCounter, out);
                        boxStack.pop();
                        continue;
                    }
                }

                dfsCollect(child, sdtStack, boxStack, boxDepthCounters, sdtCounter, out);
            } while (cur.toNextSibling());
        }
    }

    private static class SdtMeta { String alias = "", tag = ""; }
    private static SdtMeta readSdtMeta(XmlObject sdt) {
        SdtMeta m = new SdtMeta();
        try (XmlCursor cur = sdt.newCursor()) {
            if (cur.toFirstChild()) {
                do {
                    if (QN_W_SDTPr.equals(cur.getName())) {
                        try (XmlCursor c2 = cur.getObject().newCursor()) {
                            if (c2.toFirstChild()) {
                                do {
                                    QName n = c2.getName();
                                    if (QN_W_ALIAS.equals(n)) {
                                        String v = c2.getAttributeText(QN_W_VAL);
                                        if (v != null) m.alias = v;
                                    } else if (QN_W_TAG.equals(n)) {
                                        String v = c2.getAttributeText(QN_W_VAL);
                                        if (v != null) m.tag = v;
                                    }
                                } while (c2.toNextSibling());
                            }
                        }
                        break;
                    }
                } while (cur.toNextSibling());
            }
        }
        return m;
    }

    private static boolean hasAncestor(XmlObject node, QName qn) {
        try (XmlCursor c = node.newCursor()) {
            while (c.toParent()) {
                QName n = c.getName();
                if (qn.equals(n)) return true;
            }
        }
        return false;
    }

    private static String getParagraphText(XmlObject p) {
        StringBuilder sb = new StringBuilder();
        try (XmlCursor cur = p.newCursor()) {
            cur.selectPath("declare namespace w='" + NS_W + "' ./w:r");
            while (cur.toNextSelection()) {
                XmlObject r = cur.getObject();
                sb.append(getRunTextPreserveBrTab(r));
            }
        }
        return sb.toString();
    }

    private static String getRunTextPreserveBrTab(XmlObject r) {
        StringBuilder sb = new StringBuilder();
        try (XmlCursor rc = r.newCursor()) {
            if (rc.toFirstChild()) {
                do {
                    QName n = rc.getName(); if (n == null) continue;
                    String ln = n.getLocalPart();
                    if ("t".equals(ln) || "instrText".equals(ln)) {
                        String v = rc.getTextValue(); if (v != null) sb.append(v);
                    } else if ("br".equals(ln) || "cr".equals(ln)) {
                        sb.append('\n');
                    } else if ("tab".equals(ln)) {
                        sb.append('\t');
                    } else if ("softHyphen".equals(ln)) {
                        sb.append('\u00AD');
                    } else if ("noBreakHyphen".equals(ln)) {
                        sb.append('\u2011');
                    }
                } while (rc.toNextSibling());
            }
        }
        return sb.toString();
    }

    private static boolean applyChangesToPart(XmlObject root, List<XmlChange> changes) {
        if (changes == null || changes.isEmpty()) return false;

        List<XmlChange> sdtChanges = new ArrayList<>();
        List<XmlChange> tbRunChanges = new ArrayList<>();

        for (XmlChange ch : changes) {
            if ("docxSdtField".equalsIgnoreCase(ch.type)) sdtChanges.add(ch);
            else if ("docxTextBoxRun".equalsIgnoreCase(ch.type)) tbRunChanges.add(ch);
        }
        if (sdtChanges.isEmpty() && tbRunChanges.isEmpty()) return false;

        boolean[] modified = new boolean[]{false};
        Deque<Integer> sdtStack = new ArrayDeque<>();
        Deque<Integer> boxStack = new ArrayDeque<>();
        Map<Integer,Integer> boxDepthCounters = new HashMap<>();
        int[] sdtCounter = new int[]{0};

        applyDFS(root, sdtStack, boxStack, boxDepthCounters, sdtCounter, sdtChanges, tbRunChanges, modified);
        return modified[0];
    }

    private static void applyDFS(XmlObject node, Deque<Integer> sdtStack, Deque<Integer> boxStack,
                                 Map<Integer,Integer> boxDepthCounters, int[] sdtCounter,
                                 List<XmlChange> sdtChanges, List<XmlChange> tbRunChanges, boolean[] modified) {
        try (XmlCursor cur = node.newCursor()) {
            if (!cur.toFirstChild()) return;
            do {
                QName name = cur.getName();
                XmlObject child = cur.getObject();

                if (name != null) {
                    if (QN_W_SDT.equals(name)) {
                        int idx = ++sdtCounter[0];
                        sdtStack.push(idx);

                        String curSdtPath = pathString(sdtStack);
                        for (XmlChange ch : sdtChanges) {
                            if (pathEquals(ch.sdtPath, curSdtPath)) {
                                if (setSdtContentText(child, ch.newText)) modified[0] = true;
                            }
                        }

                        applyDFS(child, sdtStack, boxStack, boxDepthCounters, sdtCounter, sdtChanges, tbRunChanges, modified);
                        sdtStack.pop();
                        continue;
                    }

                    if (QN_W_TXBX_CONTENT.equals(name)) {
                        int depth = boxStack.size();
                        int next = boxDepthCounters.getOrDefault(depth, 0) + 1;
                        boxDepthCounters.put(depth, next);
                        boxStack.push(next);

                        String curSdt = pathString(sdtStack);
                        String curBox = pathString(boxStack);

                        for (XmlChange ch : tbRunChanges) {
                            if (!pathEquals(ch.sdtPath, curSdt)) continue;
                            if (!pathEquals(ch.boxPath, curBox)) continue;
                            if (ch.runOrd == null || ch.runOrd < 0) continue;
                            if (setTextBoxRunTextByOrdinal(child, ch.runOrd, ch.newText)) modified[0] = true;
                        }

                        applyDFS(child, sdtStack, boxStack, boxDepthCounters, sdtCounter, sdtChanges, tbRunChanges, modified);
                        boxStack.pop();
                        continue;
                    }
                }

                applyDFS(child, sdtStack, boxStack, boxDepthCounters, sdtCounter, sdtChanges, tbRunChanges, modified);
            } while (cur.toNextSibling());
        }
    }

    /** 文本框：若目标 run 含 <w:tab/> 则不做写入与清理 */
    private static boolean setTextBoxRunTextByOrdinal(XmlObject txbxNode, int runOrd, String text) {
        try (XmlCursor rc = txbxNode.newCursor()) {
            rc.selectPath("declare namespace w='" + NS_W + "' .//w:r");
            int idx = 0;
            while (rc.toNextSelection()) {
                XmlObject r = rc.getObject();
                if (idx == runOrd) {
                    if (runHasTabXml(r)) return true; // 保留
                    try (XmlCursor c = r.newCursor()) {
                        if (c.toFirstChild()) {
                            do {
                                QName n = c.getName(); if (n == null) continue;
                                String ln = n.getLocalPart();
                                if ("t".equals(ln) || "br".equals(ln) || "cr".equals(ln) || "instrText".equals(ln)) {
                                    c.removeXml();
                                }
                            } while (c.toNextSibling());
                        }
                    }
                    String s = (text == null) ? "" : text.replace("\r\n","\n").replace('\r','\n')
                                                .replace('\u2028','\n').replace('\u2029','\n');
                    String[] lines = s.split("\n", -1);
                    lines = trimTrailingEmpty(lines);

                    try (XmlCursor c = r.newCursor()) {
                        c.toEndToken();
                        for (int i = 0; i < lines.length; i++) {
                            if (i > 0) { c.beginElement(new QName(NS_W, "br")); c.toParent(); }
                            c.beginElement(new QName(NS_W, "t"));
                            c.insertAttributeWithValue(QN_XML_SPACE, "preserve");
                            c.insertChars(lines[i] == null ? "" : lines[i]);
                            c.toParent();
                        }
                    }
                    return true;
                }
                idx++;
            }
        } catch (Exception ignore) {}
        return false;
    }

    /** SDT：若内容包含 <w:tab/>，只覆盖现有 <w:t> 文本，不重建结构 */
    private static boolean setSdtContentText(XmlObject sdtNode, String text) {
        XmlObject sdtContent = null;
        try (XmlCursor cur = sdtNode.newCursor()) {
            if (cur.toFirstChild()) {
                do { if (QN_W_SDTCONTENT.equals(cur.getName())) { sdtContent = cur.getObject(); break; } }
                while (cur.toNextSibling());
            }
        }
        if (sdtContent == null) return false;

        if (scopeHasTab(sdtContent)) {
            return replaceAllWTInScope(sdtContent, text);
        }

        try {
            var blk = (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtContentBlock)
                    sdtContent.changeType(org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtContentBlock.type);
            if (blk != null) {
                org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr pprTpl = null;
                org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr rprTpl = null;
                if (blk.sizeOfPArray() > 0) {
                    var p0 = blk.getPArray(0);
                    pprTpl = p0.getPPr();
                    if (p0.sizeOfRArray() > 0) rprTpl = p0.getRArray(0).getRPr();
                }
                while (blk.sizeOfPArray() > 0) blk.removeP(0);

                String[] paras = splitToParagraphs(text);
                for (String para : paras) {
                    var p = blk.addNewP();
                    if (pprTpl != null) p.addNewPPr().set(pprTpl);
                    var r = p.addNewR();
                    if (rprTpl != null) r.addNewRPr().set(rprTpl);

                    String[] lines = splitToLines(para);
                    for (int i = 0; i < lines.length; i++) {
                        if (i > 0) r.addNewBr();
                        var t = r.addNewT(); t.setStringValue(lines[i]);
                        try (XmlCursor tc = t.newCursor()) { tc.setAttributeText(QN_XML_SPACE, "preserve"); } catch (Exception ignore) {}
                    }
                }
                return true;
            }
        } catch (Exception ignore) {}

        try {
            var run = (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtContentRun)
                    sdtContent.changeType(org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtContentRun.type);
            if (run != null) {
                org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr rprTpl = null;
                if (run.sizeOfRArray() > 0) rprTpl = run.getRArray(0).getRPr();
                while (run.sizeOfRArray() > 0) run.removeR(0);

                String[] parts = ANY_BREAK.split(text == null ? "" : text, -1);
                parts = trimTrailingEmpty(parts);

                var r = run.addNewR();
                if (rprTpl != null) r.addNewRPr().set(rprTpl);
                for (int i = 0; i < parts.length; i++) {
                    if (i > 0) r.addNewBr();
                    var t = r.addNewT(); t.setStringValue(parts[i]);
                    try (XmlCursor tc = t.newCursor()) { tc.setAttributeText(QN_XML_SPACE, "preserve"); } catch (Exception ignore) {}
                }
                return true;
            }
        } catch (Exception ignore) {}

        return replaceAllWTInScope(sdtContent, text);
    }

    private static boolean replaceAllWTInScope(XmlObject scope, String newText) {
        boolean found = false;
        try (XmlCursor cur = scope.newCursor()) {
            cur.selectPath("declare namespace w='" + NS_W + "' .//w:t");
            int idx = 0;
            while (cur.toNextSelection()) {
                found = true;
                try {
                    String v = cur.getAttributeText(QN_XML_SPACE);
                    if (!"preserve".equals(v)) cur.setAttributeText(QN_XML_SPACE, "preserve");
                } catch (Throwable ignore) {
                    try { cur.removeAttribute(QN_XML_SPACE); cur.setAttributeText(QN_XML_SPACE, "preserve"); }
                    catch (Throwable ignore2) {}
                }
                cur.setTextValue(idx == 0 ? (newText != null ? newText : "") : "");
                idx++;
            }
        }
        return found;
    }

    private static boolean scopeHasTab(XmlObject scope) {
        try (XmlCursor cur = scope.newCursor()) {
            cur.selectPath("declare namespace w='" + NS_W + "' .//w:tab");
            return cur.toNextSelection();
        }
    }

    private static boolean runHasTabXml(XmlObject r) {
        try (XmlCursor c = r.newCursor()) {
            c.selectPath("declare namespace w='" + NS_W + "' ./w:tab|.//w:tab");
            return c.toNextSelection();
        }
    }

    private static String aggregateParagraphSeparated(XmlObject scope, QName limit) {
        XmlObject base = scope;
        if (limit != null) {
            try (XmlCursor cur = scope.newCursor()) {
                if (cur.toFirstChild()) {
                    do { if (limit.equals(cur.getName())) { base = cur.getObject(); break; } }
                    while (cur.toNextSibling());
                }
            }
        }
        StringBuilder sb = new StringBuilder();
        try (XmlCursor c = base.newCursor()) {
            c.selectPath("declare namespace w='" + NS_W + "' .//w:p");
            boolean firstP = true;
            while (c.toNextSelection()) {
                if (!firstP) sb.append(PARA_SEP);
                firstP = false;
                sb.append(getParagraphText(c.getObject()));
            }
        }
        return sb.toString();
    }

    private static String[] trimTrailingEmpty(String[] arr) {
        int end = arr.length;
        while (end > 1 && (arr[end - 1] == null || arr[end - 1].isEmpty())) end--;
        return Arrays.copyOf(arr, end);
    }

    private static String[] splitToParagraphs(String text) {
        if (text == null) return new String[]{""};
        String s = text.replace("\r\n", "\n").replace('\r','\n').replace('\u2029','\n').replace('\u2028','\n');
        String[] paras = s.split("\n{2,}", -1);
        return trimTrailingEmpty(paras);
    }

    private static String[] splitToLines(String paragraph) {
        String[] lines = (paragraph == null ? new String[]{""} : paragraph.split("\n", -1));
        return trimTrailingEmpty(lines);
    }

    private static boolean notBlank(String s) { return s != null && !s.trim().isEmpty(); }
    private static String asString(Object v, String def) { return v == null ? def : String.valueOf(v); }
    private static boolean pathEquals(String a, String b) { return Objects.equals(a, b); }
    private static String pathString(Deque<Integer> stack) {
        if (stack.isEmpty()) return "-";
        Iterator<Integer> it = stack.descendingIterator();
        StringBuilder sb = new StringBuilder();
        while (it.hasNext()) { if (sb.length()>0) sb.append('/'); sb.append(it.next()); }
        return sb.toString();
    }

    public static void debugScanAfterWriteContainers(XWPFDocument doc) throws Exception {
        ByteArrayOutputStream bos = new ByteArrayOutputStream(); doc.write(bos); bos.flush();
        try (XWPFDocument reopened = new XWPFDocument(new ByteArrayInputStream(bos.toByteArray()))) {
            List<TextElement> els = extractWordTexts(reopened);
            int tbxRun = 0, sdt = 0;
            for (TextElement el : els) {
                if ("docxTextBoxRun".equalsIgnoreCase(el.type)) tbxRun++;
                else if ("docxSdtField".equalsIgnoreCase(el.type)) sdt++;
            }
            System.out.println("TextBoxRuns="+tbxRun+", SDT-Fields="+sdt);
        }
    }
}
