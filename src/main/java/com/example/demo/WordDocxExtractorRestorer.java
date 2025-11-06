package com.example.demo;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;

import javax.xml.namespace.QName;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.*;

/** 自包含：支持 DOCX 正文/表格 + 文本框 + SDT 的提取与回填（内存树写回） */
public class WordDocxExtractorRestorer {

    public static class TextElement {
        public final String text;
        public final String type; // "run" | "tableCell" | "docxTextBox" | "docxSdtField"
        public final Map<String, Object> position;

        public TextElement(String text, String type, Map<String, Object> position) {
            this.text = text;
            this.type = type;
            this.position = position;
        }
    }

    // 写回变化单元
    private static class XmlChange {
        final String type;       // "docxTextBox" | "docxSdtField"
        final String part;       // 通常 "/word/document.xml"
        final String sdtPath;    // 形如 "1/3/2" 或 "-"（无 SDT）
        final String boxPath;    // 形如 "1/2"   或 "-"（无 TextBox）
        final String alias;      // SDT alias，可空
        final String tag;        // SDT tag，可空
        final String kind;       // "w:txbxContent"
        final String newText;

        XmlChange(String type, String part, String sdtPath, String boxPath,
                  String alias, String tag, String kind, String newText) {
            this.type = type;
            this.part = part;
            this.sdtPath = sdtPath;
            this.boxPath = boxPath;
            this.alias = alias;
            this.tag = tag;
            this.kind = kind;
            this.newText = newText;
        }
    }

    // ======== 命名空间 & QNames ========
    private static final String NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static final QName QN_W_SDT          = new QName(NS_W, "sdt");
    private static final QName QN_W_SDTCONTENT   = new QName(NS_W, "sdtContent");
    private static final QName QN_W_SDTPr        = new QName(NS_W, "sdtPr");
    private static final QName QN_W_ALIAS        = new QName(NS_W, "alias");
    private static final QName QN_W_TAG          = new QName(NS_W, "tag");
    private static final QName QN_W_VAL          = new QName(NS_W, "val");
    private static final QName QN_W_TXBX_CONTENT = new QName(NS_W, "txbxContent");
    private static final QName QN_XML_SPACE      = new QName("http://www.w3.org/XML/1998/namespace", "space", "xml");

    // =========================
    // ======= 提  取 ==========
    // =========================

    /** 提取：正文 run/表格 + 文本框 + SDT 字段（一次性返回） */
    public static List<TextElement> extractWordTexts(XWPFDocument doc) {
        List<TextElement> elements = new ArrayList<>();

        // 1) 正文段落（run）
        for (int pIdx = 0; pIdx < doc.getParagraphs().size(); pIdx++) {
            XWPFParagraph p = doc.getParagraphs().get(pIdx);
            for (int rIdx = 0; rIdx < p.getRuns().size(); rIdx++) {
                XWPFRun run = p.getRuns().get(rIdx);
                String text = run.getText(0);
                if (notBlank(text)) {
                    Map<String, Object> pos = new HashMap<>();
                    pos.put("paragraphIndex", pIdx);
                    pos.put("runIndex", rIdx);
                    elements.add(new TextElement(text, "run", pos));
                }
            }
        }


        // 2) 表格（run 粒度；不再整块覆盖 cell）
        for (int tIdx = 0; tIdx < doc.getTables().size(); tIdx++) {
            XWPFTable t = doc.getTables().get(tIdx);
            for (int r = 0; r < t.getRows().size(); r++) {
                XWPFTableRow row = t.getRows().get(r);
                for (int c = 0; c < row.getTableCells().size(); c++) {
                    XWPFTableCell cell = row.getTableCells().get(c);
                    List<XWPFParagraph> paras = cell.getParagraphs();
                    for (int pi = 0; pi < paras.size(); pi++) {
                        XWPFParagraph p = paras.get(pi);
                        for (int ri = 0; ri < p.getRuns().size(); ri++) {
                            XWPFRun run = p.getRuns().get(ri);
                            String text = run.getText(0);
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

        // 3) /word/document.xml 内部的 文本框 & SDT 字段
        collectDocXmlContainers(doc.getDocument(), elements);

        return elements;
    }

    /** 扫描 CTDocument1（内存树）抓取 文本框 & SDT 字段 */
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
                    // 命中 SDT：提取 alias/tag、其内部文本，作为独立可回填单元
                    if (QN_W_SDT.equals(name)) {
                        int idx = ++sdtCounter[0];
                        sdtStack.push(idx);

                        SdtMeta meta = readSdtMeta(child);
                        String text = aggregateAllWT(child, QN_W_SDTCONTENT);
                        if (notBlank(meta.alias) && notBlank(text)) {
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

                        // 递归
                        dfsCollect(child, sdtStack, boxStack, boxDepthCounters, sdtCounter, out);
                        sdtStack.pop();
                        continue;
                    }

                    // 命中文本框：仅当其内部**不包含 SDT**时采集，避免容器文本与 SDT 重复
                    if (QN_W_TXBX_CONTENT.equals(name)) {
                        int depth = boxStack.size();
                        int next = boxDepthCounters.getOrDefault(depth, 0) + 1;
                        boxDepthCounters.put(depth, next);
                        boxStack.push(next);

                        boolean hasSDTInside = hasDescendant(child, QN_W_SDT);
                        if (!hasSDTInside) {
                            String txt = aggregateAllWT(child, null);
                            if (notBlank(txt)) {
                                Map<String, Object> pos = new HashMap<>();
                                pos.put("type", "docxTextBox");
                                pos.put("part", "/word/document.xml");
                                pos.put("sdtPath", pathString(sdtStack));
                                pos.put("boxPath", pathString(boxStack));
                                pos.put("kind", "w:txbxContent");
                                out.add(new TextElement(txt, "docxTextBox", pos));
                            }
                        }

                        // 递归
                        dfsCollect(child, sdtStack, boxStack, boxDepthCounters, sdtCounter, out);
                        boxStack.pop();
                        continue;
                    }
                }

                // 默认递归
                dfsCollect(child, sdtStack, boxStack, boxDepthCounters, sdtCounter, out);
            } while (cur.toNextSibling());
        }
    }

    // 读取 SDT 的 alias 与 tag
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

    // 判断是否含有某个后代元素
    private static boolean hasDescendant(XmlObject scope, QName qn) {
        try (XmlCursor cur = scope.newCursor()) {
            cur.selectPath("declare namespace w='" + NS_W + "' .//w:" + qn.getLocalPart());
            return cur.toNextSelection();
        }
    }

    // 聚合 scope 下所有 w:t 文本（若限定节点名 nonNullLimit，仅限该子树内）
    private static String aggregateAllWT(XmlObject scope, QName nonNullLimit) {
        StringBuilder sb = new StringBuilder();
        XmlObject base = scope;

        try (XmlCursor cur = scope.newCursor()) {
            if (nonNullLimit != null) {
                boolean found = false;
                if (cur.toFirstChild()) {
                    do {
                        if (nonNullLimit.equals(cur.getName())) {
                            base = cur.getObject();
                            found = true;
                            break;
                        }
                    } while (cur.toNextSibling());
                }
                if (!found) return "";
            }
        }

        try (XmlCursor c2 = base.newCursor()) {
            c2.selectPath("declare namespace w='" + NS_W + "' .//w:t");
            while (c2.toNextSelection()) {
                String v = c2.getTextValue();
                if (v != null) sb.append(v);
            }
        }
        return sb.toString();
    }

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

    private static boolean notBlank(String s) {
        return s != null && !s.trim().isEmpty();
    }

    private static String asString(Object v, String def) {
        return v == null ? def : String.valueOf(v);
    }

    private static String[] splitPreserveEmpty(String s) {
        if (s == null) return new String[]{""};
        // 按换行拆分，保留空白行
        return s.split("\\R", -1);
    }

    private static boolean pathEquals(String a, String b) {
        if (Objects.equals(a, b)) return true;
        return false;
    }

    // =========================
    // ======= 回  填 ==========
    // =========================

    /** 回填：正文 run/表格 + 文本框 + SDT（仅在内存 CT 上修改，随后对 doc.write(...) 序列化即可） */
    public static void restoreWordTexts(XWPFDocument doc, List<TextElement> elements, List<String> translated) {
        // 1) run & tableRun：就地改字，完全保留 rPr
        for (int i = 0; i < elements.size(); i++) {
            TextElement el = elements.get(i);
            String t = translated.get(i);

            if ("run".equals(el.type)) {
                int p = (Integer) el.position.get("paragraphIndex");
                int r = (Integer) el.position.get("runIndex");
                XWPFRun run = doc.getParagraphs().get(p).getRuns().get(r);
                run.setText(t, 0);

            } else if ("tableRun".equals(el.type)) {
                int ti = (Integer) el.position.get("tableIndex");
                int ri = (Integer) el.position.get("rowIndex");
                int ci = (Integer) el.position.get("cellIndex");
                int pi = (Integer) el.position.get("paraInCell");
                int rui = (Integer) el.position.get("runInPara");
                XWPFRun run = doc.getTables().get(ti)
                        .getRow(ri).getCell(ci)
                        .getParagraphArray(pi).getRuns().get(rui);
                run.setText(t, 0);
            }
        }

        // 2) 收集对 /word/document.xml 的变更（文本框 & SDT 字段）
        Map<String, List<XmlChange>> changesByPart = new HashMap<>();
        for (int i = 0; i < elements.size(); i++) {
            TextElement el = elements.get(i);
            String typ = el.type == null ? "" : el.type;
            if (!typ.equalsIgnoreCase("docxTextBox") && !typ.equalsIgnoreCase("textBox")
             && !typ.equalsIgnoreCase("docxSdtField") && !typ.equalsIgnoreCase("sdtField")) {
                continue;
            }
            String part    = asString(el.position.get("part"), "/word/document.xml");
            String sdtPath = asString(el.position.get("sdtPath"), "-");
            String boxPath = asString(el.position.get("boxPath"), "-");
            String alias   = asString(el.position.get("alias"), "");
            String tag     = asString(el.position.get("tag"), "");
            String kind    = asString(el.position.get("kind"), "w:txbxContent");
            String newTxt  = translated.get(i);

            XmlChange ch = new XmlChange(typ, part, sdtPath, boxPath, alias, tag, kind, newTxt);
            changesByPart.computeIfAbsent(part, k -> new ArrayList<>()).add(ch);
        }

        // 3) 仅在内存 CT 上应用变更（避免包级写回破坏）
        List<XmlChange> docChanges = changesByPart.get("/word/document.xml");
        if (docChanges != null && !docChanges.isEmpty()) {
            applyChangesToPart(doc.getDocument(), docChanges);
            // boolean modified = applyChangesToPart(doc.getDocument(), docChanges);
            // System.out.println("[restore] in-memory applyChangesToPart modified=" + modified);
        }

        // 4) 同步封面属性（Title/Subtitle/Author），避免模板覆盖
        syncCoverProperties(doc, elements, translated);
    }

    private static void syncCoverProperties(XWPFDocument doc, List<TextElement> elements, List<String> translated) {
        try {
            org.apache.poi.ooxml.POIXMLProperties props = doc.getProperties();
            org.apache.poi.ooxml.POIXMLProperties.CoreProperties core = props.getCoreProperties();
            org.apache.poi.ooxml.POIXMLProperties.CustomProperties custom = props.getCustomProperties();

            for (int i = 0; i < elements.size(); i++) {
                TextElement el = elements.get(i);
                String type = el.type == null ? "" : el.type;
                if (!type.equalsIgnoreCase("docxSdtField") && !type.equalsIgnoreCase("sdtField")) continue;

                String alias = asString(el.position.get("alias"), "");
                String txt   = translated.get(i);

                if ("Title".equalsIgnoreCase(alias)) {
                    try { core.setTitle(txt); } catch (Throwable ignore) {}
                } else if ("Author".equalsIgnoreCase(alias)) {
                    try { core.setCreator(txt); } catch (Throwable ignore) {}
                } else if ("Subtitle".equalsIgnoreCase(alias)) {
                    // 兼容不同 POI 版本的 Subject 写法（优先 setSubjectProperty，其次 setSubject）
                    setCoreSubjectCompat(core, txt);

                    // 尝试写 Custom 属性 "Subtitle" —— 只在不存在时新增；不做更新，避免版本差异
                    try {
                        if (custom != null) {
                            // 有些版本：getProperty(String) 可用
                            try {
                                java.lang.reflect.Method getProp = custom.getClass().getMethod("getProperty", String.class);
                                Object cp = getProp.invoke(custom, "Subtitle");
                                if (cp == null) {
                                    java.lang.reflect.Method add = custom.getClass().getMethod("addProperty", String.class, String.class);
                                    add.invoke(custom, "Subtitle", txt);
                                }
                            } catch (Throwable ignore) {
                                // 版本不支持 getProperty：忽略即可（封面依据 Subject 已同步）
                            }
                        }
                    } catch (Throwable ignore) {}
                }
            }
        } catch (Throwable ignore) {}
    }

    // CoreProperties 的 Subject 版本兼容（反射）
    private static void setCoreSubjectCompat(org.apache.poi.ooxml.POIXMLProperties.CoreProperties core, String txt) {
        try {
            // 新一些的 POI
            java.lang.reflect.Method m = core.getClass().getMethod("setSubjectProperty", String.class);
            m.invoke(core, txt);
            return;
        } catch (Throwable ignore) { /* fallthrough */ }
        try {
            // 有些版本的方法名叫 setSubject
            java.lang.reflect.Method m = core.getClass().getMethod("setSubject", String.class);
            m.invoke(core, txt);
        } catch (Throwable ignore) { /* 放弃：封面仍可通过 SDT 内容本身改好 */ }
    }

    // ===== 内存 CT 树写回（文本框 & SDT） =====
    private static boolean applyChangesToPart(XmlObject root, List<XmlChange> changes) {
        if (changes == null || changes.isEmpty()) return false;

        List<XmlChange> tbChanges = new ArrayList<>();
        List<XmlChange> sdtChanges = new ArrayList<>();
        for (XmlChange ch : changes) {
            String t = ch.type == null ? "" : ch.type;
            if (t.equalsIgnoreCase("docxTextBox") || t.equalsIgnoreCase("textBox")) tbChanges.add(ch);
            else if (t.equalsIgnoreCase("docxSdtField") || t.equalsIgnoreCase("sdtField")) sdtChanges.add(ch);
        }
        // System.out.println("[restore] applyChangesToPart: tb=" + tbChanges.size() + ", sdt=" + sdtChanges.size());
        if (tbChanges.isEmpty() && sdtChanges.isEmpty()) return false;

        boolean[] modified = new boolean[]{false};
        Deque<Integer> sdtStack = new ArrayDeque<>();
        Deque<Integer> boxStack = new ArrayDeque<>();
        Map<Integer,Integer> boxDepthCounters = new HashMap<>();
        int[] sdtCounter = new int[]{0};

        applyDFS(root, sdtStack, boxStack, boxDepthCounters, sdtCounter, tbChanges, sdtChanges, modified);
        // System.out.println("[restore] applyChangesToPart: modified=" + modified[0]);
        return modified[0];
    }

    private static void applyDFS(XmlObject node, Deque<Integer> sdtStack, Deque<Integer> boxStack, Map<Integer,Integer> boxDepthCounters,
                                 int[] sdtCounter, List<XmlChange> tbChanges, List<XmlChange> sdtChanges, boolean[] modified) {
        try (XmlCursor cur = node.newCursor()) {
            if (!cur.toFirstChild()) return;
            do {
                QName name = cur.getName();
                XmlObject child = cur.getObject();

                if (name != null) {
                    // SDT：alias/tag 为主，路径尽量约束
                    if (QN_W_SDT.equals(name)) {
                        int idx = ++sdtCounter[0];
                        sdtStack.push(idx);

                        SdtMeta meta = readSdtMeta(child);
                        String curSdtPath = pathString(sdtStack);
                        String curBoxPath = pathString(boxStack);

                        for (XmlChange ch : sdtChanges) {
                            // alias/tag 至少一个匹配
                            boolean aliasOK = notBlank(ch.alias) && ch.alias.equals(meta.alias);
                            boolean tagOK   = notBlank(ch.tag)   && ch.tag.equals(meta.tag);
                            if (!aliasOK && !tagOK) continue;

                            // 路径弱约束（不一致时也允许，但若 change 指定了 boxPath，尽量匹配）
                            boolean boxPathOK = pathEquals(ch.boxPath, curBoxPath) || "-".equals(ch.boxPath);
                            if (!boxPathOK && notBlank(ch.boxPath) && !"-".equals(ch.boxPath)) {
                                continue;
                            }

                            boolean ok = setSdtContentText(child, ch.newText);
                            // System.out.println("[write-SDT] @" + curSdtPath + "/" + curBoxPath + " ok=" + ok);
                            if (ok) modified[0] = true;
                        }

                        applyDFS(child, sdtStack, boxStack, boxDepthCounters, sdtCounter, tbChanges, sdtChanges, modified);
                        sdtStack.pop();
                        continue;
                    }

                    // 文本框：路径精确匹配
                    if (QN_W_TXBX_CONTENT.equals(name)) {
                        int depth = boxStack.size();
                        int next = boxDepthCounters.getOrDefault(depth, 0) + 1;
                        boxDepthCounters.put(depth, next);
                        boxStack.push(next);

                        String curSdt = pathString(sdtStack);
                        String curBox = pathString(boxStack);

                        for (XmlChange ch : tbChanges) {
                            if (!pathEquals(ch.sdtPath, curSdt)) continue;
                            if (!pathEquals(ch.boxPath, curBox)) continue;

                            boolean ok = setTextBoxText(child, ch.newText);
                            // System.out.println("[write-TBX] @" + curSdt + "/" + curBox + " ok=" + ok);
                            if (ok) modified[0] = true;
                        }

                        applyDFS(child, sdtStack, boxStack, boxDepthCounters, sdtCounter, tbChanges, sdtChanges, modified);
                        boxStack.pop();
                        continue;
                    }
                }

                applyDFS(child, sdtStack, boxStack, boxDepthCounters, sdtCounter, tbChanges, sdtChanges, modified);
            } while (cur.toNextSibling());
        }
    }

    // === 公用：就地替换，完整保留样式（优先策略） ===
    private static boolean replaceTextPreserveFormatting(XmlObject scope, String text) {
        // 找到第一个 w:r 作为“样式承载”
        org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR firstRun = null;
        try (XmlCursor cur = scope.newCursor()) {
            cur.selectPath("declare namespace w='" + NS_W + "' .//w:r");
            if (cur.toNextSelection()) {
                firstRun = (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR) cur.getObject();
            }
        } catch (Exception ignore) {}

        if (firstRun == null) {
            // 没有任何 run：退化为“就地替换所有 w:t”的兜底
            return replaceAllWTInScope(scope, text);
        }

        // 清掉 firstRun 内现有 w:t
        try (XmlCursor rc = firstRun.newCursor()) {
            rc.selectPath("declare namespace w='" + NS_W + "' ./w:t");
            while (rc.toNextSelection()) rc.removeXml();
        } catch (Exception ignore) {}

        // 写入文本（处理换行 → w:br）
        String[] lines = splitPreserveEmpty(text);
        for (int i = 0; i < lines.length; i++) {
            if (i > 0) firstRun.addNewBr();
            var t = firstRun.addNewT();
            t.setStringValue(lines[i]);
            try (XmlCursor tc = t.newCursor()) { tc.setAttributeText(QN_XML_SPACE, "preserve"); } catch (Exception ignore) {}
        }

        // 其它 w:t 清空，避免残留
        try (XmlCursor cur = scope.newCursor()) {
            cur.selectPath("declare namespace w='" + NS_W + "' .//w:t");
            boolean skippedFirst = false;
            while (cur.toNextSelection()) {
                if (!skippedFirst) { skippedFirst = true; continue; }
                cur.setTextValue("");
            }
        } catch (Exception ignore) {}

        return true;
    }

    // === 写入：文本框（优先就地；必要时重建 + 拷贝样式） ===
    private static boolean setTextBoxText(XmlObject txbxNode, String text) {
        // 优先：就地改字，完全保留 pPr/rPr
        if (replaceTextPreserveFormatting(txbxNode, text)) return true;

        // 备选：重建，但复制第一段/第一 run 的样式为模板
        try {
            var tx = (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTxbxContent)
                    txbxNode.changeType(org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTxbxContent.type);
            if (tx != null) {
                org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr pprTpl = null;
                org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr rprTpl = null;
                if (tx.sizeOfPArray() > 0) {
                    var p0 = tx.getPArray(0);
                    pprTpl = p0.getPPr();
                    if (p0.sizeOfRArray() > 0) rprTpl = p0.getRArray(0).getRPr();
                }

                while (tx.sizeOfPArray() > 0) tx.removeP(0);
                String[] lines = splitPreserveEmpty(text);
                for (String line : lines) {
                    var p = tx.addNewP();
                    if (pprTpl != null) p.addNewPPr().set(pprTpl);
                    var r = p.addNewR();
                    if (rprTpl != null) r.addNewRPr().set(rprTpl);
                    var t = r.addNewT(); t.setStringValue(line);
                    try (XmlCursor tc = t.newCursor()) { tc.setAttributeText(QN_XML_SPACE, "preserve"); } catch (Exception ignore) {}
                }
                return true;
            }
        } catch (Exception ignore) {}
        return replaceAllWTInScope(txbxNode, text);
    }
    
    // === 写入：SDT（优先就地；必要时重建 + 拷贝样式） ===
    private static boolean setSdtContentText(XmlObject sdtNode, String text) {
        XmlObject sdtContent = null;
        try (XmlCursor cur = sdtNode.newCursor()) {
            if (cur.toFirstChild()) {
                do {
                    if (QN_W_SDTCONTENT.equals(cur.getName())) { sdtContent = cur.getObject(); break; }
                } while (cur.toNextSibling());
            }
        }
        if (sdtContent == null) return false;

        // 优先：就地改字（完整保留 pPr/rPr）
        if (replaceTextPreserveFormatting(sdtContent, text)) return true;

        // 备选1：块级重建（复制模板样式）
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
                String[] lines = splitPreserveEmpty(text);
                for (String line : lines) {
                    var p = blk.addNewP();
                    if (pprTpl != null) p.addNewPPr().set(pprTpl);
                    var r = p.addNewR();
                    if (rprTpl != null) r.addNewRPr().set(rprTpl);
                    var t = r.addNewT(); t.setStringValue(line);
                    try (XmlCursor tc = t.newCursor()) { tc.setAttributeText(QN_XML_SPACE, "preserve"); } catch (Exception ignore) {}
                }
                return true;
            }
        } catch (Exception ignore) {}

        // 备选2：行级重建（复制模板样式）
        try {
            var run = (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtContentRun)
                    sdtContent.changeType(org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtContentRun.type);
            if (run != null) {
                org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr rprTpl = null;
                if (run.sizeOfRArray() > 0) rprTpl = run.getRArray(0).getRPr();

                while (run.sizeOfRArray() > 0) run.removeR(0);
                String[] lines = splitPreserveEmpty(text);
                for (int i = 0; i < lines.length; i++) {
                    String line = lines[i];
                    var r = run.addNewR();
                    if (rprTpl != null) r.addNewRPr().set(rprTpl);
                    var t = r.addNewT(); t.setStringValue(line);
                    try (XmlCursor tc = t.newCursor()) { tc.setAttributeText(QN_XML_SPACE, "preserve"); } catch (Exception ignore) {}
                    if (i < lines.length - 1) r.addNewBr();
                }
                return true;
            }
        } catch (Exception ignore) {}

        // 兜底
        return replaceAllWTInScope(sdtContent, text);
    }
    
    /** 在给定作用域就地替换所有 <w:t>：第一个写 newText，其余清空；先设 xml:space="preserve"，再写文本 */
    private static boolean replaceAllWTInScope(XmlObject scope, String newText) {
        boolean found = false;
        try (XmlCursor cur = scope.newCursor()) {
            cur.selectPath("declare namespace w='" + NS_W + "' .//w:t");
            int idx = 0;
            while (cur.toNextSelection()) {
                found = true;
                // 先属性后文本
                try {
                    String v = cur.getAttributeText(QN_XML_SPACE);
                    if (!"preserve".equals(v)) {
                        cur.setAttributeText(QN_XML_SPACE, "preserve");
                    }
                } catch (Throwable ignore) {
                    try {
                        cur.removeAttribute(QN_XML_SPACE);
                        cur.setAttributeText(QN_XML_SPACE, "preserve");
                    } catch (Throwable ignore2) {}
                }
                cur.setTextValue(idx == 0 ? (newText != null ? newText : "") : "");
                idx++;
            }
        }
        return found;
    }

    // ========= 调试：落盘重开扫描，验证与 Word 所见一致 =========
    public static void debugScanAfterWriteContainers(XWPFDocument doc) throws Exception {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        doc.write(bos);
        bos.flush();

        try (XWPFDocument reopened = new XWPFDocument(new ByteArrayInputStream(bos.toByteArray()))) {
            System.out.println("==== [REOPEN 扫描：容器快照] ====");

            List<TextElement> els = extractWordTexts(reopened);
            int tbx = 0, sdt = 0;

            for (TextElement el : els) {
                if ("docxTextBox".equalsIgnoreCase(el.type)) {
                    tbx++;
                    String sdtPath = String.valueOf(el.position.getOrDefault("sdtPath", "-"));
                    String boxPath = String.valueOf(el.position.getOrDefault("boxPath", "-"));
                    System.out.println("[TextBox] SDT=" + sdtPath + " Box=" + boxPath + " → \"" + el.text + "\"");
                } else if ("docxSdtField".equalsIgnoreCase(el.type)) {
                    sdt++;
                    String sdtPath = String.valueOf(el.position.getOrDefault("sdtPath", "-"));
                    String boxPath = String.valueOf(el.position.getOrDefault("boxPath", "-"));
                    String alias   = String.valueOf(el.position.getOrDefault("alias", ""));
                    String tag     = String.valueOf(el.position.getOrDefault("tag", ""));
                    System.out.println("[SDT-Field] SDT=" + sdtPath + " Box=" + boxPath +
                            " alias=\"" + alias + "\" tag=\"" + tag + "\" → \"" + el.text + "\"");
                }
            }

            System.out.println("-- 汇总: TextBox=" + tbx + ", SDT-Fields=" + sdt + " --");
            System.out.println("==== [END REOPEN] ====");
        }
    }
}
