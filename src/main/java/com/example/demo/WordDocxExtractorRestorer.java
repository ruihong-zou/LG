package com.example.demo;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;

import javax.xml.namespace.QName;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.*;
import java.util.function.Function;
import java.util.regex.Pattern;

/** 支持 DOCX：正文/表格按“格式变化”分段提取与原位回填；文本框 run 与（文本框外）SDT 的精确写回；提供空白保护编解码 */
public class WordDocxExtractorRestorer {

    // ===== 数据模型 =====
    public static class TextElement {
        public final String text;
        public final String type;              // paraSeg | cellSeg | docxTextBoxRun | docxSdtField
        public final Map<String, Object> position;

        public TextElement(String text, String type, Map<String, Object> position) {
            this.text = text; this.type = type; this.position = position;
        }
    }

    /** 文本框与 SDT 的写回单元 */
    private static class XmlChange {
        final String type;     // "docxTextBoxRun" | "docxSdtField"
        final String part;     // "/word/document.xml"
        final String sdtPath;  // SDT 路径（DFS 计数）
        final String boxPath;  // 文本框路径（DFS 计数）
        final String alias;    // SDT alias
        final String tag;      // SDT tag
        final String kind;     // 容器类型（调试）
        final String newText;  // 写回文本
        final Integer runOrd;  // 文本框中的 run 顺序号

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

    private static final String PARA_SEP = "\u2029";
    private static final Pattern ANY_BREAK = Pattern.compile("\r\n|\r|\n|\u2028|\u2029|\u000B|\u000C|\u0085");

    // ====================== 提取 ======================
    /** 提取文档中的文本元素，正文/表格按“格式变化”分段，文本框与（文本框外）SDT 单独处理 */
    public static List<TextElement> extractWordTexts(XWPFDocument doc) {
        List<TextElement> elements = new ArrayList<>();

        // 1) 正文段落 → paraSeg
        List<XWPFParagraph> paras = doc.getParagraphs();
        for (int pIdx = 0; pIdx < paras.size(); pIdx++) {
            XWPFParagraph p = paras.get(pIdx);
            List<FormatChangeSegmenter.Segment> segs = FormatChangeSegmenter.segmentParagraph(p);
            for (FormatChangeSegmenter.Segment seg : segs) {
                if (seg.text == null) continue;
                Map<String, Object> pos = new HashMap<>();
                pos.put("paragraphIndex", pIdx);
                pos.put("runFrom", seg.runStartIdx);
                pos.put("runTo", seg.runEndIdx);
                elements.add(new TextElement(seg.text, "paraSeg", pos));
            }
        }

        // 2) 表格单元格 → cellSeg
        for (int tIdx = 0; tIdx < doc.getTables().size(); tIdx++) {
            XWPFTable t = doc.getTables().get(tIdx);
            for (int r = 0; r < t.getRows().size(); r++) {
                XWPFTableRow row = t.getRows().get(r);
                for (int c = 0; c < row.getTableCells().size(); c++) {
                    XWPFTableCell cell = row.getTableCells().get(c);
                    List<XWPFParagraph> ps = cell.getParagraphs();
                    for (int pi = 0; pi < ps.size(); pi++) {
                        XWPFParagraph p = ps.get(pi);
                        List<FormatChangeSegmenter.Segment> segs = FormatChangeSegmenter.segmentParagraph(p);
                        for (FormatChangeSegmenter.Segment seg : segs) {
                            if (seg.text == null) continue;
                            Map<String, Object> pos = new HashMap<>();
                            pos.put("tableIndex", tIdx);
                            pos.put("rowIndex", r);
                            pos.put("cellIndex", c);
                            pos.put("paraInCell", pi);
                            pos.put("runFrom", seg.runStartIdx);
                            pos.put("runTo", seg.runEndIdx);
                            elements.add(new TextElement(seg.text, "cellSeg", pos));
                        }
                    }
                }
            }
        }

        // 3) 文本框 run + SDT（仅文本框外）
        collectDocXmlContainers(doc.getDocument(), elements);

        return elements;
    }

    /** 扫描 document.xml，枚举文本框内 .//w:r（runOrd）与“文本框外”的 SDT */
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
                    // SDT（仅文本框外）
                    if (QN_W_SDT.equals(name)) {
                        int idx = ++sdtCounter[0];
                        sdtStack.push(idx);

                        boolean insideTextBox = hasAncestor(child, QN_W_TXBX_CONTENT);
                        if (!insideTextBox) {
                            SdtMeta meta = readSdtMeta(child);
                            String text = aggregateParagraphSeparated(child, QN_W_SDTCONTENT);
                            if (text != null) {
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

                    // 文本框内容
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
                                String txt = getRunTextPreserveBrTab(rObj);
                                if (txt != null) {
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
    private static class SdtMeta { String alias = ""; String tag = ""; }
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
                if (qn.equals(c.getName())) return true;
            }
        }
        return false;
    }

    // ===== 读取 run 文本（保留换行与制表） =====
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

    // ====================== 空白保护（翻译前后） ======================
    /** 空白保护编码与解码工具 */
    public static final class WhitespaceGuard {

        private static final Pattern LEADING = Pattern.compile("^[ \\t]+");
        private static final Pattern TRAILING = Pattern.compile("[ \\t]+$");
        private static final Pattern MULTI_SPACES = Pattern.compile(" {2,}");
        private static final Pattern PH_TOKEN = Pattern.compile("⟦s(\\d+)⟧");

        /** 单条文本的行级空白信息 */
        public static final class LineMeta {
            public final String leading;  // 行首原始空白（空格/Tab）
            public final String trailing; // 行尾原始空白（空格/Tab）
            public LineMeta(String l, String t) { this.leading = l; this.trailing = t; }
        }

        /** 批量的元数据容器 */
        public static final class Batch {
            public final List<List<LineMeta>> metas; // 每条文本按行的空白信息
            public Batch(List<List<LineMeta>> metas) { this.metas = metas; }
        }

        /** 对一批文本做空白编码，返回编码后的文本与元数据 */
        public static Encoded encodeBatch(List<String> texts) {
            List<String> encoded = new ArrayList<>(texts.size());
            List<List<LineMeta>> metas = new ArrayList<>(texts.size());
            for (String s : texts) {
                if (s == null) s = "";
                String[] lines = s.split("\r\n|\r|\n",-1);
                List<LineMeta> lm = new ArrayList<>(lines.length);
                for (int i = 0; i < lines.length; i++) {
                    String line = lines[i];

                    String lead = "";
                    var m1 = LEADING.matcher(line);
                    if (m1.find()) { lead = m1.group(); line = line.substring(lead.length()); }

                    String trail = "";
                    var m2 = TRAILING.matcher(line);
                    if (m2.find()) { trail = m2.group(); line = line.substring(0, line.length()-trail.length()); }

                    String protectedInside = MULTI_SPACES.matcher(line).replaceAll(match -> {
                        int n = match.group().length();
                        return "⟦s" + n + "⟧";
                    });

                    lines[i] = protectedInside;
                    lm.add(new LineMeta(lead, trail));
                }
                encoded.add(String.join("\n", lines));
                metas.add(lm);
            }
            return new Encoded(encoded, new Batch(metas));
        }

        /** 将编码后的翻译结果按元数据还原空白 */
        public static List<String> decodeBatch(List<String> translatedEncoded, Batch batchMeta) {
            List<String> out = new ArrayList<>(translatedEncoded.size());
            for (int i = 0; i < translatedEncoded.size(); i++) {
                String s = translatedEncoded.get(i);
                List<LineMeta> lm = batchMeta.metas.get(i);
                String[] lines = s.split("\r\n|\r|\n",-1);
                // 若翻译端改了行数，以最短为准，剩余行按空字符串处理
                int L = Math.max(lines.length, lm.size());
                String[] norm = new String[L];
                for (int k = 0; k < L; k++) {
                    String line = (k < lines.length) ? lines[k] : "";
                    line = PH_TOKEN.matcher(line).replaceAll(m -> " ".repeat(Integer.parseInt(m.group(1))));
                    String leading = (k < lm.size()) ? lm.get(k).leading : "";
                    String trailing = (k < lm.size()) ? lm.get(k).trailing : "";
                    norm[k] = leading + line + trailing;
                }
                out.add(String.join("\n", norm));
            }
            return out;
        }

        /** 编码批次的返回体 */
        public static final class Encoded {
            public final List<String> encoded;
            public final Batch meta;
            public Encoded(List<String> encoded, Batch meta) { this.encoded = encoded; this.meta = meta; }
        }
    }

    /** 使用空白保护的翻译管线：传入批量翻译函数，返回解码后的译文 */
    public static List<String> translateWithWhitespaceGuard(
            List<String> plainTexts,
            Function<List<String>, List<String>> batchTranslator) {
        var enc = WhitespaceGuard.encodeBatch(plainTexts);
        List<String> tr = batchTranslator.apply(enc.encoded);
        return WhitespaceGuard.decodeBatch(tr, enc.meta);
    }

    // ====================== 回填 ======================
    /** 将翻译后的文本写回文档，正文/表格按段写回，文本框与 SDT 按路径/序号写回 */
    public static void restoreWordTexts(XWPFDocument doc, List<TextElement> elements, List<String> translated) {
        if (elements == null || translated == null || elements.size() != translated.size()) {
            throw new IllegalArgumentException("elements 与 translated 数量不一致");
        }

        // 1) 正文/表格 段级回填
        for (int i = 0; i < elements.size(); i++) {
            TextElement el = elements.get(i);
            String t = translated.get(i);
            if ("paraSeg".equals(el.type)) {
                int pIdx = (Integer) el.position.get("paragraphIndex");
                int from = (Integer) el.position.get("runFrom");
                int to   = (Integer) el.position.get("runTo");
                XWPFParagraph p = doc.getParagraphs().get(pIdx);
                SegmentRestorer.restoreSegmentInParagraph(p, from, to, t);

            } else if ("cellSeg".equals(el.type)) {
                int ti = (Integer) el.position.get("tableIndex");
                int ri = (Integer) el.position.get("rowIndex");
                int ci = (Integer) el.position.get("cellIndex");
                int pi = (Integer) el.position.get("paraInCell");
                int from = (Integer) el.position.get("runFrom");
                int to   = (Integer) el.position.get("runTo");
                XWPFTableCell cell = doc.getTables().get(ti).getRow(ri).getCell(ci);
                SegmentRestorer.restoreSegmentInCell(cell, pi, from, to, t);
            }
        }

        // 2) 文本框 run 与 SDT 写回
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

    /** 应用文本框与 SDT 的写回到指定 XML part */
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

    /** 深度遍历应用文本框与 SDT 写回 */
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
                            if (!pathEquals(ch.sdtPath, curSdtPath)) continue;
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

    // ====== 文本框：按 runOrd 写回（支持换行与制表） ======
    private static boolean setTextBoxRunTextByOrdinal(XmlObject txbxNode, int runOrd, String text) {
        try (XmlCursor rc = txbxNode.newCursor()) {
            rc.selectPath("declare namespace w='" + NS_W + "' .//w:r");
            int idx = 0;
            while (rc.toNextSelection()) {
                if (idx == runOrd) {
                    XmlObject r = rc.getObject();
                    // 清理文本子节点，保留 rPr
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
                    String s = (text == null) ? "" : text.replace("\r\n","\n").replace('\r','\n')
                                                .replace('\u2028','\n').replace('\u2029','\n');
                    String[] lines = s.split("\n", -1);
                    lines = trimTrailingEmpty(lines);

                    try (XmlCursor c = r.newCursor()) {
                        c.toEndToken();
                        for (int i = 0; i < lines.length; i++) {
                            if (i > 0) { c.beginElement(new QName(NS_W, "br")); c.toParent(); }
                            String[] parts = lines[i].split("\t", -1);
                            for (int j = 0; j < parts.length; j++) {
                                c.beginElement(new QName(NS_W, "t"));
                                c.insertAttributeWithValue(QN_XML_SPACE, "preserve");
                                c.insertChars(parts[j] == null ? "" : parts[j]);
                                c.toParent();
                                if (j < parts.length - 1) { c.beginElement(new QName(NS_W, "tab")); c.toParent(); }
                            }
                        }
                    }
                    return true;
                }
                idx++;
            }
        } catch (Exception ignore) {}
        return false;
    }

    // ====== SDT：写回文本（支持多段、多行、制表） ======
    private static boolean setSdtContentText(XmlObject sdtNode, String text) {
        XmlObject sdtContent = null;
        try (XmlCursor cur = sdtNode.newCursor()) {
            if (cur.toFirstChild()) {
                do { if (QN_W_SDTCONTENT.equals(cur.getName())) { sdtContent = cur.getObject(); break; } }
                while (cur.toNextSibling());
            }
        }
        if (sdtContent == null) return false;

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
                        appendLineWithTabs(r, lines[i]);
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
                    appendLineWithTabs(r, parts[i]);
                }
                return true;
            }
        } catch (Exception ignore) {}

        return replaceAllWTInScope(sdtContent, text);
    }

    /** 为一个 CT-Run 追加一行文本：按 \t 切分为多个 <w:t xml:space="preserve">，段间插入 <w:tab/> */
    private static void appendLineWithTabs(org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR r, String line) {
        String[] parts = line.split("\t", -1);
        for (int j = 0; j < parts.length; j++) {
            var t = r.addNewT();
            t.setStringValue(parts[j] == null ? "" : parts[j]);
            try (XmlCursor tc = t.newCursor()) { tc.setAttributeText(QN_XML_SPACE, "preserve"); } catch (Exception ignore) {}
            if (j < parts.length - 1) r.addNewTab();
        }
    }

    // ===== 工具：通用替换/聚合/字符串处理 =====
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

    private static String asString(Object v, String def) { return v == null ? def : String.valueOf(v); }
    private static boolean pathEquals(String a, String b) { return Objects.equals(a, b); }
    private static String pathString(Deque<Integer> stack) {
        if (stack.isEmpty()) return "-";
        Iterator<Integer> it = stack.descendingIterator();
        StringBuilder sb = new StringBuilder();
        while (it.hasNext()) { if (sb.length() > 0) sb.append('/'); sb.append(it.next()); }
        return sb.toString();
    }

    // ====================== 调试（可选） ======================
    /** 重新打开文档并扫描文本框与 SDT 提取结果 */
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
                            " alias=" + asString(el.position.get("alias"),"") + " tag=" +
                            asString(el.position.get("tag"),"") + " → \"" + el.text + "\"");
                }
            }
            System.out.println("-- 汇总: TextBoxRuns=" + tbxRun + ", SDT-Fields=" + sdt + " --");
            System.out.println("==== [END] ====");
        }
    }

}
