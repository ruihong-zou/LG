package com.example.demo;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import java.lang.reflect.*;
import java.util.*;

/**
 * .doc（HWPF）版：正文/表格/文本框(TextBoxes Story)/页眉/脚注等的提取与回填 + 自检
 */
public class WordDocExtractorRestorer {

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

    // ========= 1) 提取 =========
    public static List<TextElement> extractWordTexts(HWPFDocument doc) {
        List<TextElement> elements = new ArrayList<>();
        Range range = doc.getRange();

        // 1.1 主文档：普通段落（跳过表格内段落）
        for (int pIdx = 0; pIdx < range.numParagraphs(); pIdx++) {
            Paragraph para = range.getParagraph(pIdx);
            if (para.isInTable()) continue;
            for (int rIdx = 0; rIdx < para.numCharacterRuns(); rIdx++) {
                CharacterRun run = para.getCharacterRun(rIdx);
                String txt = run.text();
                if (txt != null && !txt.trim().isEmpty()) {
                    Map<String,Object> pos = new HashMap<>();
                    pos.put("paragraphIndex", pIdx);
                    pos.put("runIndex", rIdx);
                    elements.add(new TextElement(txt, "run", pos));
                }
            }
        }

        // 1.2 主文档：表格 run
        List<Table> tables = new ArrayList<>();
        TableIterator tit = new TableIterator(range);
        while (tit.hasNext()) tables.add(tit.next());

        for (int tIdx = 0; tIdx < tables.size(); tIdx++) {
            Table table = tables.get(tIdx);
            for (int rowIdx = 0; rowIdx < table.numRows(); rowIdx++) {
                TableRow row = table.getRow(rowIdx);
                for (int cellIdx = 0; cellIdx < row.numCells(); cellIdx++) {
                    TableCell cell = row.getCell(cellIdx);
                    for (int cp = 0; cp < cell.numParagraphs(); cp++) {
                        Paragraph cellPara = cell.getParagraph(cp);
                        for (int cr = 0; cr < cellPara.numCharacterRuns(); cr++) {
                            CharacterRun run = cellPara.getCharacterRun(cr);
                            String txt = run.text();
                            if (txt != null) txt = txt.replaceAll("\\p{Cntrl}", "");
                            if (txt != null && !txt.trim().isEmpty()) {
                                Map<String,Object> pos = new HashMap<>();
                                pos.put("tableIndex", tIdx);
                                pos.put("rowIndex", rowIdx);
                                pos.put("cellIndex", cellIdx);
                                pos.put("cellParaIndex", cp);
                                pos.put("cellRunIndex", cr);
                                elements.add(new TextElement(txt, "tableCellRun", pos));
                            }
                        }
                    }
                }
            }
        }

        // 1.3 文本框故事流
        Range tb = tryGetTextboxRange(doc); // 见下方辅助方法
        if (tb != null) {
            for (int p = 0; p < tb.numParagraphs(); p++) {
                Paragraph para = tb.getParagraph(p);
                for (int r = 0; r < para.numCharacterRuns(); r++) {
                    CharacterRun run = para.getCharacterRun(r);
                    String txt = run.text();
                    if (txt != null) txt = txt.replace("\r", "").replaceAll("\\p{Cntrl}", "");
                    if (txt != null && !txt.trim().isEmpty()) {
                        Map<String,Object> pos = new HashMap<>();
                        pos.put("tbParaIndex", p);
                        pos.put("tbRunIndex", r);
                        elements.add(new TextElement(txt, "tbRun", pos));
                    }
                }
            }
        }

        return elements;
    }

    // ========= 2) 回写 =========
    public static void restoreWordTexts(HWPFDocument doc, List<TextElement> elements, List<String> translatedTexts) {
        Range docRange = doc.getRange();

        // 缓存主文档的表格（和你原先一致）
        List<Table> tables = new ArrayList<>();
        TableIterator tit = new TableIterator(docRange);
        while (tit.hasNext()) tables.add(tit.next());

        // 文本框故事流（关键新增）
        Range tbRange = tryGetTextboxRange(doc);

        // 倒序替换，避免索引错乱（保持你原先策略）
        for (int idx = elements.size() - 1; idx >= 0; idx--) {
            TextElement el = elements.get(idx);
            String newRaw = translatedTexts.get(idx);
            if (newRaw == null) newRaw = "";
            newRaw = newRaw.replace("\n", "");

            CharacterRun run = null;
            String oldFull = null;
            boolean hasCR = false;

            switch (el.type) {
                case "run": {
                    Integer pI = (Integer) el.position.get("paragraphIndex");
                    Integer rI = (Integer) el.position.get("runIndex");
                    if (pI == null || rI == null) break;
                    if (pI < 0 || pI >= docRange.numParagraphs()) break;
                    Paragraph para = docRange.getParagraph(pI);
                    if (rI < 0 || rI >= para.numCharacterRuns()) break;
                    run = para.getCharacterRun(rI);
                    oldFull = run.text();
                    if (oldFull == null) oldFull = "";
                    hasCR = oldFull.endsWith("\r");
                    break;
                }
                case "tableCellRun": {
                    Integer tI  = (Integer) el.position.get("tableIndex");
                    Integer rI  = (Integer) el.position.get("rowIndex");
                    Integer cI  = (Integer) el.position.get("cellIndex");
                    Integer cpI = (Integer) el.position.get("cellParaIndex");
                    Integer crI = (Integer) el.position.get("cellRunIndex");
                    if (tI == null || rI == null || cI == null || cpI == null || crI == null) break;
                    if (tI < 0 || tI >= tables.size()) break;
                    TableCell cell = tables.get(tI).getRow(rI).getCell(cI);
                    Paragraph cellPara = cell.getParagraph(cpI);
                    if (crI < 0 || crI >= cellPara.numCharacterRuns()) break;
                    run = cellPara.getCharacterRun(crI);
                    oldFull = run.text();
                    if (oldFull == null) oldFull = "";
                    oldFull = oldFull.replaceAll("\\p{Cntrl}", "");
                    hasCR = oldFull.endsWith("\r");
                    break;
                }
                case "tbRun": { // 文本框 story 的回写：等长 + 控制位骨架保留 + 等长 token 两步替换
                    if (tbRange == null) break;

                    Integer pI = (Integer) el.position.get("tbParaIndex");
                    Integer rI = (Integer) el.position.get("tbRunIndex");
                    if (pI == null || rI == null) break;
                    if (pI < 0 || pI >= tbRange.numParagraphs()) break;
                    Paragraph p = tbRange.getParagraph(pI);
                    if (rI < 0 || rI >= p.numCharacterRuns()) break;

                    run = p.getCharacterRun(rI);
                    oldFull = run.text();
                    if (oldFull == null) oldFull = "";
                    if (oldFull.isEmpty()) break;                // 空串直接跳过，避免卡死

                    // 保留尾部 \r 情况
                    hasCR = oldFull.endsWith("\r");
                    String oldCore = hasCR ? oldFull.substring(0, oldFull.length() - 1) : oldFull;

                    // 新文本（不带换行；控制位用骨架保留逻辑处理）
                    String newCore = translatedTexts.get(idx);
                    if (newCore == null) newCore = "";
                    newCore = newCore.replace("\r", "").replace("\n", "");

                    // —— 骨架回填：控制/方向/制表/空格原样保留，仅可见位依序填入译文，长度==oldCore.length()
                    StringBuilder filled = new StringBuilder(oldCore.length());
                    int srcPos = 0;
                    for (int pos = 0; pos < oldCore.length(); pos++) {
                        char ch = oldCore.charAt(pos);
                        boolean isControl =
                                ch < 0x20 || ch == 0x7F ||
                                (ch >= 0x200E && ch <= 0x200F) ||     // LRM/RLM
                                (ch >= 0x202A && ch <= 0x202E);       // LRE/RLE/PDF/LRO/RLO
                        boolean preserveWhitespace = (ch == '\t' || ch == ' ');

                        if (isControl || preserveWhitespace) {
                            filled.append(ch);                        // 这些位置绝对不能动
                        } else {
                            if (srcPos < newCore.length()) filled.append(newCore.charAt(srcPos++));
                            else filled.append(' ');                  // 不够则补空格，保持等长
                        }
                    }
                    String newFull = hasCR ? (filled.toString() + "\r") : filled.toString();

                    // 若完全一致就不必替换，避免无谓扫描
                    if (oldFull.equals(newFull)) break;

                    // —— 两步替换（token 必须与 oldFull 等长，并在控制位处“原样”保留）
                    String token = buildSameLenTokenPreservingControls(oldFull);
                    run.replaceText(oldFull, token);    // 中间态长度不变
                    run.replaceText(token,   newFull);  // 最终态长度也不变
                    break;
                }
            }

            if (run != null && oldFull != null) {
                // 生成安全 token（避免 old/new 重叠导致 replaceText 异常）
                String tokenCore = generateSafeToken(oldFull, newRaw);
                String token     = tokenCore + (hasCR ? "\r" : "");
                String newFull   = newRaw    + (hasCR ? "\r" : "");

                run.replaceText(oldFull, token);
                run.replaceText(token,   newFull);
            }
        }
    }

    // ========= 3) 获取文本框故事流 Range（带回退） =========
    private static Range tryGetTextboxRange(HWPFDocument doc) {
        try {
            Range r = doc.getMainTextboxRange(); // POI 4.x/5.x
            if (r != null) return r;
        } catch (Throwable ignore) {}
        try {
            // 某些分支实现保留了旧名
            java.lang.reflect.Method m = HWPFDocument.class.getMethod("getTextboxesRange");
            Object ret = m.invoke(doc);
            if (ret instanceof Range) return (Range) ret;
        } catch (Throwable ignore) {}
        return null;
    }

    private static boolean isControlOrDirMark(char ch) {
        return ch < 0x20 || ch == 0x7F ||
            (ch >= 0x200E && ch <= 0x200F) ||
            (ch >= 0x202A && ch <= 0x202E);
    }

    /** 生成“与 oldFull 等长”的 token；控制位/空白原样保留，其余用私有区字符填充，几乎不可能撞车 */
    private static String buildSameLenTokenPreservingControls(String oldFull) {
        if (oldFull == null) return "";
        StringBuilder sb = new StringBuilder(oldFull.length());
        for (int i = 0; i < oldFull.length(); i++) {
            char ch = oldFull.charAt(i);
            if (isControlOrDirMark(ch) || ch == '\t' || ch == ' ') {
                sb.append(ch); // 这些位置保持原样（包括末尾 \r）
            } else {
                // 用私有区 U+E000 起的字符，占位但不易与正文冲突
                sb.append((char) (0xE000 + (i % 32)));
            }
        }
        return sb.toString();
    }

    // ================== 自检 / 调试 ==================
    public static void debugSnapshot(HWPFDocument doc) {
        System.out.println("==== [.DOC 容器快照] ====");

        // 1) 主文档
        try {
            Range main = doc.getRange();
            snapshotOne(main, "[main]", null);
        } catch (Throwable t) {
            System.out.println("[main] 访问失败: " + t.getClass().getSimpleName() + ": " + t.getMessage());
        }

        // 2) 文本框故事流（关键：用详细模式逐 run 打印）
        try {
            Range tb = tryGetRange(doc, "getMainTextboxRange");
            if (tb == null) tb = tryGetRange(doc, "getTextboxesRange"); // 旧方法名兜底
            if (tb != null) {
                snapshotTextboxVerbose(tb, "[textbox]");
            } else {
                System.out.println("[textbox] (no range via getMainTextboxRange/getTextboxesRange)");
            }
        } catch (Throwable t) {
            System.out.println("[textbox] 访问失败: " + t.getClass().getSimpleName() + ": " + t.getMessage());
        }

        // 3) 页眉/页脚
        snapshotOptional(doc, "getHeaderStoryRange", "[header]");
        snapshotOptional(doc, "getFooterStoryRange", "[footer]");

        // 4) 脚注、尾注、批注（延用你原来的签名）
        snapshotOptional(doc, "getFootnoteRange", "[footnote]");
        snapshotOptional(doc, "getEndnoteRange",  "[endnote]");
        snapshotOptional(doc, "getCommentsRange", "[comments]");

        System.out.println("==== [END .DOC 容器快照] ====");
    }

    private static void snapshotTextboxVerbose(Range r, String label) {
        if (r == null) {
            System.out.println(label + " (null)");
            return;
        }
        int paraCnt = r.numParagraphs();
        int runCnt  = 0;
        int tableCnt = 0, tableRuns = 0;

        System.out.println(label + " 段落=" + paraCnt);
        // 逐段逐 run 打印
        for (int p = 0; p < paraCnt; p++) {
            Paragraph para = r.getParagraph(p);
            int nRuns = para.numCharacterRuns();
            for (int i = 0; i < nRuns; i++) {
                CharacterRun cr = para.getCharacterRun(i);
                String t = clean(cr.text());   // 仅清理 \r 和控制符，保持文本
                if (!t.isEmpty()) {
                    System.out.println(String.format("  (p=%d r=%d) \"%s\"", p, i, shorten(t, 200)));
                    runCnt++;
                }
            }
        }

        // 表格（极少见，仍统计）
        try {
            TableIterator it = new TableIterator(r);
            while (it.hasNext()) {
                tableCnt++;
                Table tbl = it.next();
                for (int row = 0; row < tbl.numRows(); row++) {
                    TableRow tr = tbl.getRow(row);
                    for (int c = 0; c < tr.numCells(); c++) {
                        TableCell cell = tr.getCell(c);
                        for (int p = 0; p < cell.numParagraphs(); p++) {
                            Paragraph cp = cell.getParagraph(p);
                            for (int k = 0; k < cp.numCharacterRuns(); k++) {
                                String t = clean(cp.getCharacterRun(k).text());
                                if (!t.isEmpty()) {
                                    tableRuns++;
                                    System.out.println(String.format(
                                        "  (table row=%d col=%d p=%d r=%d) \"%s\"",
                                        row, c, p, k, shorten(t, 200)
                                    ));
                                }
                            }
                        }
                    }
                }
            }
        } catch (Throwable ignore) {}

        // 汇总
        System.out.println(String.format(
            "%s 段落runs=%d, 表格数=%d, 表格runs=%d",
            label.replace("[", "[[").replace("]", "]]"), runCnt, tableCnt, tableRuns
        ));

        // 拼接所有 run，进一步目测是否漏字
        StringBuilder all = new StringBuilder();
        for (int p = 0; p < paraCnt; p++) {
            Paragraph para = r.getParagraph(p);
            for (int i = 0; i < para.numCharacterRuns(); i++) {
                all.append(clean(para.getCharacterRun(i).text())).append(" | ");
            }
        }
        System.out.println(label + " all-runs: " + shorten(all.toString(), 1000));
    }

    private static void snapshotOptional(HWPFDocument doc, String getterName, String storyKey) {
        Range r = tryGetRange(doc, getterName);
        if (r != null) snapshotOne(r, storyKey, null);
        else System.out.println("[" + storyKey + "] (no range via " + getterName + ")");
    }

    private static String clean(String s) {
        if (s == null) return "";
        // HWPF 的 run.text() 常带 \r 作为段落结束标记，这里只做快照显示用
        s = s.replace("\r", "");
        // 去掉不可见控制符，避免把域控制符打印出来
        return s.replaceAll("\\p{Cntrl}", "").trim();
    }

    private static String shorten(String s, int max) {
        if (s == null) return "";
        return s.length() <= max ? s : s.substring(0, max) + "…";
    }
    
    private static Range tryGetRange(HWPFDocument doc, String... names) {
        for (String n : names) {
            try {
                Method m = HWPFDocument.class.getMethod(n);
                Object r = m.invoke(doc);
                if (r instanceof Range) return (Range) r;
            } catch (Throwable ignore) {}
        }
        return null;
    }

    private static void snapshotOne(Range range, String storyKey, Map<String,Object> hint) {
        if (range == null) return;
        int paraRuns = 0, cellRuns = 0, tbls = 0;

        for (int p = 0; p < range.numParagraphs(); p++) {
            Paragraph para = range.getParagraph(p);
            if (para.isInTable()) continue;
            for (int r = 0; r < para.numCharacterRuns(); r++) {
                String t = sanitize(para.getCharacterRun(r).text());
                if (!isBlank(t)) paraRuns++;
            }
        }

        TableIterator it = new TableIterator(range);
        while (it.hasNext()) {
            tbls++;
            Table tbl = it.next();
            for (int i = 0; i < tbl.numRows(); i++) {
                TableRow row = tbl.getRow(i);
                for (int j = 0; j < row.numCells(); j++) {
                    TableCell cell = row.getCell(j);
                    for (int cp = 0; cp < cell.numParagraphs(); cp++) {
                        Paragraph para = cell.getParagraph(cp);
                        for (int cr = 0; cr < para.numCharacterRuns(); cr++) {
                            String t = sanitize(para.getCharacterRun(cr).text());
                            if (!isBlank(t)) cellRuns++;
                        }
                    }
                }
            }
        }

        String extra = (hint != null && hint.containsKey("textboxIndex"))
                ? " (tb#" + hint.get("textboxIndex") + ")"
                : "";
        System.out.println("[" + storyKey + extra + "] 段落runs=" + paraRuns + ", 表格数=" + tbls + ", 表格runs=" + cellRuns);

        int printed = 0;
        for (int p = 0; p < range.numParagraphs() && printed < 3; p++) {
            Paragraph para = range.getParagraph(p);
            if (para.isInTable()) continue;
            for (int r = 0; r < para.numCharacterRuns() && printed < 3; r++) {
                String t = sanitize(para.getCharacterRun(r).text());
                if (!isBlank(t)) {
                    System.out.println("  (sample) \"" + t + "\"");
                    printed++;
                }
            }
        }
    }

    // ==== 通用工具 ====
    private static String sanitize(String txt) { return (txt == null) ? null : txt.replaceAll("\\p{Cntrl}", ""); }
    
    private static boolean isBlank(String s) { return s == null || s.trim().isEmpty(); }
    
    private static String generateSafeToken(String oldFull, String newFull) {
        Set<Character> forbid = new HashSet<>();
        if (oldFull != null) for (char c : oldFull.toCharArray()) forbid.add(c);
        if (newFull != null) for (char c : newFull.toCharArray()) forbid.add(c);
        String base = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
        List<Character> allowed = new ArrayList<>();
        for (char c : base.toCharArray()) if (!forbid.contains(c)) allowed.add(c);
        if (allowed.isEmpty()) {
            int code = 0xE000 + new java.util.Random().nextInt(0x1000);
            return new String(Character.toChars(code));
        }
        StringBuilder sb = new StringBuilder(8);
        java.util.Random rnd = new java.util.Random();
        for (int i = 0; i < 8; i++) sb.append(allowed.get(rnd.nextInt(allowed.size())));
        return sb.toString();
    }
}
