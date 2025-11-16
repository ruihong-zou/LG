// File: src/main/java/com/example/demo/MergePolicy.java
package com.example.demo;

/** 控制“仅因格式变化分段”的松紧度 */
public final class MergePolicy {
    public final double fontSizeTolerancePt;     // 字号容差（pt）
    public final int charSpacingTolerance;       // 字间距容差（1/100 pt）
    public final boolean ignoreFontFamilyNullVsExplicit; // null ↔ 显式字体 视同
    public final boolean ignoreColorAutoVsNull;  // auto ↔ null 视同
    public final boolean ignoreAllColorDiff;     // 忽略颜色差异
    public final boolean ignoreHighlightDiff;    // 忽略高亮差异
    public final boolean hardOnBoldDiff;         // 粗体差异是否强断开
    public final boolean hardOnItalicDiff;       // 斜体差异是否强断开
    public final boolean hardOnUnderlineDiff;    // 下划线差异是否强断开
    public final boolean hardOnStrikeDiff;       // 删除线差异是否强断开
    public final boolean hardOnFontFamilyDiff;   // 字体族差异是否强断开

    private MergePolicy(Builder b) {
        this.fontSizeTolerancePt = b.fontSizeTolerancePt;
        this.charSpacingTolerance = b.charSpacingTolerance;
        this.ignoreFontFamilyNullVsExplicit = b.ignoreFontFamilyNullVsExplicit;
        this.ignoreColorAutoVsNull = b.ignoreColorAutoVsNull;
        this.ignoreAllColorDiff = b.ignoreAllColorDiff;
        this.ignoreHighlightDiff = b.ignoreHighlightDiff;
        this.hardOnBoldDiff = b.hardOnBoldDiff;
        this.hardOnItalicDiff = b.hardOnItalicDiff;
        this.hardOnUnderlineDiff = b.hardOnUnderlineDiff;
        this.hardOnStrikeDiff = b.hardOnStrikeDiff;
        this.hardOnFontFamilyDiff = b.hardOnFontFamilyDiff;
    }

    /** 宽松合并策略：适度忽略字号/字距/auto 颜色/高亮差异，保留粗斜下划线为强边界 */
    public static MergePolicy loose() {
        return new Builder()
            .fontSizeTolerancePt(1.0)
            .charSpacingTolerance(100)
            .ignoreFontFamilyNullVsExplicit(true)
            .ignoreColorAutoVsNull(true)
            .ignoreAllColorDiff(false)
            .ignoreHighlightDiff(true)
            .hardOnBoldDiff(true)
            .hardOnItalicDiff(true)
            .hardOnUnderlineDiff(true)
            .hardOnStrikeDiff(false)
            .hardOnFontFamilyDiff(false)
            .build();
    }

    public static final class Builder {
        private double fontSizeTolerancePt = 1.0;
        private int charSpacingTolerance = 100;
        private boolean ignoreFontFamilyNullVsExplicit = true;
        private boolean ignoreColorAutoVsNull = true;
        private boolean ignoreAllColorDiff = false;
        private boolean ignoreHighlightDiff = true;
        private boolean hardOnBoldDiff = true;
        private boolean hardOnItalicDiff = true;
        private boolean hardOnUnderlineDiff = true;
        private boolean hardOnStrikeDiff = false;
        private boolean hardOnFontFamilyDiff = false;

        public Builder fontSizeTolerancePt(double v){ this.fontSizeTolerancePt=v; return this; }
        public Builder charSpacingTolerance(int v){ this.charSpacingTolerance=v; return this; }
        public Builder ignoreFontFamilyNullVsExplicit(boolean v){ this.ignoreFontFamilyNullVsExplicit=v; return this; }
        public Builder ignoreColorAutoVsNull(boolean v){ this.ignoreColorAutoVsNull=v; return this; }
        public Builder ignoreAllColorDiff(boolean v){ this.ignoreAllColorDiff=v; return this; }
        public Builder ignoreHighlightDiff(boolean v){ this.ignoreHighlightDiff=v; return this; }
        public Builder hardOnBoldDiff(boolean v){ this.hardOnBoldDiff=v; return this; }
        public Builder hardOnItalicDiff(boolean v){ this.hardOnItalicDiff=v; return this; }
        public Builder hardOnUnderlineDiff(boolean v){ this.hardOnUnderlineDiff=v; return this; }
        public Builder hardOnStrikeDiff(boolean v){ this.hardOnStrikeDiff=v; return this; }
        public Builder hardOnFontFamilyDiff(boolean v){ this.hardOnFontFamilyDiff=v; return this; }
        public MergePolicy build(){ return new MergePolicy(this); }
    }
}
