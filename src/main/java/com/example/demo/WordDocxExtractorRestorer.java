package com.example.demo;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;

import javax.xml.namespace.QName;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.*;
import java.util.regex.Pattern;

/** 支持 DOCX 正文/表格 run + 文本框 run(原位) + SDT（文本框外）的提取与回填 */
public class WordDocxExtractorRestorer {

    // ===== 数据模型 =====
    public static class TextElement {
        public final String text;
        public final String type; // "run" | "tableRun" | "docxTextBoxRun" | "docxSdtField"
        public final Map<String, Object> position;

        public TextElement(String text, String type, Map<String, Object> position) {
            this.text = text;
            this.type = type;
            this.position = position;
        }
    }

    /** 写回变化单元 */
    private static class XmlChange {
        final String type;   // "docxTextBoxRun" | "docxSdtField"
        final String part;   // "/word/document.xml"
        final String sdtPath;  // DFS 计数路径，如 "2/1"；无则 "-"
        final String boxPath;  // 文本框路径，如 "1/3"；无则 "-"
        final String alias;    // SDT alias，可空
        final String tag;      // SDT tag，可空
        final String kind;     // 容器类型（调试用）
        final String newText;

        // 文本框 run 的定位（在该文本框内 .//w:r 的顺序号）
        final Integer runOrd;

        XmlChange(String type, String part, String sdtPath, String boxPath,
                  String alias, String tag, String kind, String newText, Integer runOrd) {
            this.type = type; this.part = part; this.sdtPath = sdtPath; this.boxPath = boxPath;
            this.alias = alias; this.tag = tag; this.kind = kind; this.newText = newText; this.runOrd = runOrd;
        }
    }

    // ===== 常量 / 命名空间 =====
    private static final String NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static final QName QN_W_SDT          = new QName(NS_W, "sdt");
    private static final QName QN_W_SDTCONTENT   = new QName(NS_W, "sdtContent");
    private static final QName QN_W_SDTPr        = new QName(NS_W, "sdtPr");
    private static final QName QN_W_ALIAS        = new QName(NS_W, "alias");
    private static final QName QN_W_TAG          = new QName(NS_W, "tag");
    private static final QName QN_W_VAL          = new QName(NS_W, "val");
    private static final QName QN_W_TXBX_CONTENT = new QName(NS_W, "txbxContent");
    private static final QName QN_XML_SPACE      = new QName("http://www.w3.org/XML/1998/namespace", "space", "xml");

    private static final String PARA_SEP = "\u2029"; // 段落分隔符（用于 SDT 聚合）
    private static final Pattern ANY_BREAK = Pattern.compile("\r\n|\r|\n|\u2028|\u2029|\u000B|\u000C|\u0085");

    // ===== 提取 =====
    public static List<TextElement> extractWordTexts(XWPFDocument doc) {
        List<TextElement> elements = new ArrayList<>();

        // 正文 run
        List<XWPFParagraph> paras = doc.getParagraphs();
        for (int pIdx = 0; pIdx < paras.size(); pIdx++) {
            XWPFParagraph p = paras.get(pIdx);
            List<XWPFRun> runs = p.getRuns();
            for (int rIdx = 0; rIdx < runs.size(); rIdx++) {
                String text = runs.get(rIdx).getText(0);
                if (notBlank(text)) {
                    Map<String, Object> pos = new HashMap<>();
                    pos.put("paragraphIndex", pIdx);
                    pos.put("runIndex", rIdx);
                    elements.add(new TextElement(text, "run", pos));
                }
            }
        }

        // 表格 run
        for (int tIdx = 0; tIdx < doc.getTables().size(); tIdx++) {
            XWPFTable t = doc.getTables().get(tIdx);
            for (int r = 0; r < t.getRows().size(); r++) {
                XWPFTableRow row = t.getRows().get(r);
                for (int c = 0; c < row.getTableCells().size(); c++) {
                    XWPFTableCell cell = row.getTableCells().get(c);
                    List<XWPFParagraph> ps = cell.getParagraphs();
                    for (int pi = 0; pi < ps.size(); pi++) {
                        XWPFParagraph p = ps.get(pi);
                        List<XWPFRun> rs = p.getRuns();
                        for (int ri = 0; ri < rs.size(); ri++) {
                            String text = rs.get(ri).getText(0);
                            if (notBlank(text)) {
                                Map<String, Object> pos = new HashMap<>();
                                pos.put("tableIndex", tIdx);
                                pos.put("rowIndex", r);
                                pos.put("cellIndex", c);
                                pos.put("paraInCell", pi);
                                pos.put("runInPara", ri);
                                elements.add(new TextElement(text, "tableRun", pos));
                            }
                        }
                    }
                }
            }
        }

        // 文本框 run + SDT（仅文本框外的 SDT）
        collectDocXmlContainers(doc.getDocument(), elements);

        return elements;
    }

    /** 扫描 document.xml：枚举文本框内的所有 .//w:r（runOrd），以及文本框外层的 SDT */
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
                    // SDT：仅当不位于文本框内时，才聚合为 docxSdtField（避免与文本框 run 重复）
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
                                pos.put("alias", meta.alias); // 允许为空
                                pos.put("tag", meta.tag);     // 允许为空
                                pos.put("kind", "sdt");
                                out.add(new TextElement(text, "docxSdtField", pos));
                            }
                        }

                        dfsCollect(child, sdtStack, boxStack, boxDepthCounters, sdtCounter, out);
                        sdtStack.pop();
                        continue;
                    }

                    // 文本框：按 run 顺序号收集（包含 SDT 内部的 run）
                    if (QN_W_TXBX_CONTENT.equals(name)) {
                        int depth = boxStack.size();
                        int next = boxDepthCounters.getOrDefault(depth, 0) + 1;
                        boxDepthCounters.put(depth, next);
                        boxStack.push(next);

                        // 统计该文本框内 .//w:r 的顺序，并收集文本
                        try (XmlCursor rc = child.newCursor()) {
                            rc.selectPath("declare namespace w='" + NS_W + "' .//w:r");
                            int runOrd = 0;
                            while (rc.toNextSelection()) {
                                XmlObject rObj = rc.getObject();
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

    // ===== SDT 元数据 / 祖先检查 =====
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

    // ===== 段/运行文本读取（保留 br/tab） =====
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

    // ===== 回填 =====
    public static void restoreWordTexts(XWPFDocument doc, List<TextElement> elements, List<String> translated) {
        // 1) 正文/表格 run 原位回填
        for (int i = 0; i < elements.size(); i++) {
            TextElement el = elements.get(i);
            String t = translated.get(i);
            if ("run".equals(el.type)) {
                int p = (Integer) el.position.get("paragraphIndex");
                int r = (Integer) el.position.get("runIndex");
                setRunTextSmart(doc.getParagraphs().get(p).getRuns().get(r), t);
            } else if ("tableRun".equals(el.type)) {
                int ti = (Integer) el.position.get("tableIndex");
                int ri = (Integer) el.position.get("rowIndex");
                int ci = (Integer) el.position.get("cellIndex");
                int pi = (Integer) el.position.get("paraInCell");
                int rui = (Integer) el.position.get("runInPara");
                XWPFRun run = doc.getTables().get(ti).getRow(ri).getCell(ci)
                                   .getParagraphArray(pi).getRuns().get(rui);
                setRunTextSmart(run, t);
            }
        }

        // 2) 收集文本框 run & SDT（文本框外）
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

        // 3) 应用到内存 CT
        List<XmlChange> docChanges = changesByPart.get("/word/document.xml");
        if (docChanges != null && !docChanges.isEmpty()) {
            applyChangesToPart(doc.getDocument(), docChanges);
        }

        // 4) 同步封面（可选）
        syncCoverProperties(doc, elements, translated);
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

                        // 仅处理“文本框外”的 SDT（我们提取阶段就已过滤）
                        String curSdtPath = pathString(sdtStack);
                        String curBoxPath = pathString(boxStack);

                        for (XmlChange ch : sdtChanges) {
                            boolean sdtPathOK = pathEquals(ch.sdtPath, curSdtPath);
                            if (!sdtPathOK) continue;
                            boolean ok = setSdtContentText(child, ch.newText);
                            if (ok) modified[0] = true;
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
                            boolean ok = setTextBoxRunTextByOrdinal(child, ch.runOrd, ch.newText);
                            if (ok) modified[0] = true;
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

    // ===== 写入：文本框 run（按 runOrd 命中） =====
    private static boolean setTextBoxRunTextByOrdinal(XmlObject txbxNode, int runOrd, String text) {
        try (XmlCursor rc = txbxNode.newCursor()) {
            rc.selectPath("declare namespace w='" + NS_W + "' .//w:r");
            int idx = 0;
            while (rc.toNextSelection()) {
                if (idx == runOrd) {
                    XmlObject r = rc.getObject();
                    // 清除 r 下的 t/br/tab，保留 rPr
                    try (XmlCursor c = r.newCursor()) {
                        if (c.toFirstChild()) {
                            do {
                                QName n = c.getName(); if (n == null) continue;
                                String ln = n.getLocalPart();
                                if ("t".equals(ln) || "br".equals(ln) || "cr".equals(ln) || "tab".equals(ln) || "instrText".equals(ln)) {
                                    c.removeXml();
                                }
                            } while (c.toNextSibling());
                        }
                    }
                    // 写入（把 \r/\u2028/\u2029 统一为 \n；多行→多 br）
                    String s = (text == null) ? "" : text.replace("\r\n","\n").replace('\r','\n')
                                                .replace('\u2028','\n').replace('\u2029','\n');
                    String[] lines = s.split("\n", -1);
                    lines = trimTrailingEmpty(lines);

                    // 在末尾追加 t/br
                    try (XmlCursor c = r.newCursor()) {
                        c.toEndToken();
                        for (int i = 0; i < lines.length; i++) {
                            if (i > 0) {
                                c.beginElement(new QName(NS_W, "br"));
                                c.toParent();
                            }
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

    // ===== 写入：SDT =====
    private static boolean setSdtContentText(XmlObject sdtNode, String text) {
        XmlObject sdtContent = null;
        try (XmlCursor cur = sdtNode.newCursor()) {
            if (cur.toFirstChild()) {
                do { if (QN_W_SDTCONTENT.equals(cur.getName())) { sdtContent = cur.getObject(); break; } }
                while (cur.toNextSibling());
            }
        }
        if (sdtContent == null) return false;

        // 区块型：<w:p>...
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

        // 行型：折算为单 run + 多 br
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

    // ===== 工具：通用替换、聚合、段落操作 =====
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
        String s = text.replace("\r\n", "\n").replace('\r','\n')
                       .replace('\u2029','\n').replace('\u2028','\n');
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
        while (it.hasNext()) {
            if (sb.length() > 0) sb.append('/');
            sb.append(it.next());
        }
        return sb.toString();
    }

    // ===== 直接写 body run 的工具（正文/表格） =====
    private static void setRunTextSmart(XWPFRun run, String text) {
        if (text == null) text = "";
        String s = text.replace("\r\n","\n").replace('\r','\n').replace('\u2028','\n');
        boolean needParagraphs = s.contains("\u2029") || s.contains("\n\n");

        IRunBody parent = run.getParent();
        if (!(parent instanceof XWPFParagraph)) {
            writeLinesToRun(run, s.split("\n", -1));
            return;
        }
        XWPFParagraph p = (XWPFParagraph) parent;

        var pprTpl = p.getCTP().getPPr();
        var rprTpl = run.getCTR().getRPr();

        if (!needParagraphs) {
            writeLinesToRun(run, s.split("\n", -1));
            return;
        }

        String[] paras = trimTrailingEmpty(s.replace('\u2029','\n').split("\n{2,}", -1));

        clearParagraphRuns(p);
        XWPFRun first = p.createRun();
        applyTplStyles(p, pprTpl, first, rprTpl);
        writeLinesToRun(first, splitToLines(paras[0]));

        for (int i = 1; i < paras.length; i++) {
            XWPFParagraph np = insertParagraphAfter(p);
            clearParagraphRuns(np);
            XWPFRun r = np.createRun();
            applyTplStyles(np, pprTpl, r, rprTpl);
            writeLinesToRun(r, splitToLines(paras[i]));
            p = np;
        }
    }

    private static XWPFParagraph insertParagraphAfter(XWPFParagraph p) {
        IBody body = p.getBody();
        try (XmlCursor cursor = p.getCTP().newCursor()) {
            cursor.toEndToken(); cursor.toNextToken();
            if (body instanceof XWPFDocument) return ((XWPFDocument) body).insertNewParagraph(cursor);
            if (body instanceof XWPFTableCell) return ((XWPFTableCell) body).insertNewParagraph(cursor);
        } catch (Throwable ignore) {}
        if (body instanceof XWPFTableCell) return ((XWPFTableCell) body).addParagraph();
        if (body instanceof XWPFDocument) return ((XWPFDocument) body).createParagraph();
        return p;
    }

    private static void applyTplStyles(XWPFParagraph target,
                                       org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr pprTpl,
                                       XWPFRun run,
                                       org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr rprTpl) {
        if (pprTpl != null) {
            var tp = target.getCTP().isSetPPr() ? target.getCTP().getPPr() : target.getCTP().addNewPPr();
            tp.set(pprTpl);
        }
        if (rprTpl != null && run != null) {
            var rp = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
            rp.set(rprTpl);
        }
    }

    private static void clearParagraphRuns(XWPFParagraph p) { while (p.getRuns().size() > 0) p.removeRun(0); }

    private static void writeLinesToRun(XWPFRun run, String[] lines) {
        lines = trimTrailingEmpty(lines);
        var ctr = run.getCTR();
        try (XmlCursor rc = ctr.newCursor()) {
            rc.selectPath("declare namespace w='" + NS_W + "' ./w:t");
            while (rc.toNextSelection()) rc.removeXml();
        } catch (Exception ignore) {}
        for (int i = 0; i < lines.length; i++) {
            if (i > 0) ctr.addNewBr();
            var t = ctr.addNewT();
            t.setStringValue(lines[i] == null ? "" : lines[i]);
            try (XmlCursor tc = t.newCursor()) { tc.setAttributeText(QN_XML_SPACE, "preserve"); } catch (Exception ignore) {}
        }
    }

    // ===== 调试：重开验证 =====
    public static void debugScanAfterWriteContainers(XWPFDocument doc) throws Exception {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        doc.write(bos); bos.flush();

        try (XWPFDocument reopened = new XWPFDocument(new ByteArrayInputStream(bos.toByteArray()))) {
            System.out.println("==== [REOPEN 扫描] ====");
            List<TextElement> els = extractWordTexts(reopened);
            int tbxRun = 0, sdt = 0;
            for (TextElement el : els) {
                if ("docxTextBoxRun".equalsIgnoreCase(el.type)) {
                    tbxRun++;
                    System.out.println("[TextBoxRun] box=" + el.position.get("boxPath") +
                            " sdt=" + el.position.get("sdtPath") +
                            " ord=" + el.position.get("runOrd") +
                            " → \"" + el.text + "\"");
                } else if ("docxSdtField".equalsIgnoreCase(el.type)) {
                    sdt++;
                    System.out.println("[SDT] sdt=" + el.position.get("sdtPath") +
                            " box=" + el.position.get("boxPath") +
                            " alias=\"" + asString(el.position.get("alias"),"") + "\" tag=\"" +
                            asString(el.position.get("tag"),"") + "\" → \"" + el.text + "\"");
                }
            }
            System.out.println("-- 汇总: TextBoxRuns=" + tbxRun + ", SDT-Fields=" + sdt + " --");
            System.out.println("==== [END] ====");
        }
    }

    // ===== 封面属性（可选）=====
    private static void syncCoverProperties(XWPFDocument doc, List<TextElement> elements, List<String> translated) {
        try {
            var props = doc.getProperties();
            var core = props.getCoreProperties();
            var custom = props.getCustomProperties();

            for (int i = 0; i < elements.size(); i++) {
                TextElement el = elements.get(i);
                if (!"docxSdtField".equalsIgnoreCase(el.type)) continue;
                String alias = asString(el.position.get("alias"), "");
                String txt   = translated.get(i);
                if ("Title".equalsIgnoreCase(alias)) {
                    try { core.setTitle(txt); } catch (Throwable ignore) {}
                } else if ("Author".equalsIgnoreCase(alias)) {
                    try { core.setCreator(txt); } catch (Throwable ignore) {}
                } else if ("Subtitle".equalsIgnoreCase(alias)) {
                    setCoreSubjectCompat(core, txt);
                    try {
                        if (custom != null) {
                            try {
                                var getProp = custom.getClass().getMethod("getProperty", String.class);
                                Object cp = getProp.invoke(custom, "Subtitle");
                                if (cp == null) {
                                    var add = custom.getClass().getMethod("addProperty", String.class, String.class);
                                    add.invoke(custom, "Subtitle", txt);
                                }
                            } catch (Throwable ignore) {}
                        }
                    } catch (Throwable ignore) {}
                }
            }
        } catch (Throwable ignore) {}
    }

    private static void setCoreSubjectCompat(org.apache.poi.ooxml.POIXMLProperties.CoreProperties core, String txt) {
        try { core.getClass().getMethod("setSubjectProperty", String.class).invoke(core, txt); return; }
        catch (Throwable ignore) {}
        try { core.getClass().getMethod("setSubject", String.class).invoke(core, txt); }
        catch (Throwable ignore) {}
    }
}
