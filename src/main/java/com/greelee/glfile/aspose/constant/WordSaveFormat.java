package com.greelee.glfile.aspose.constant;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/3/30
 * @describe: 类型格式
 */
public enum WordSaveFormat {

    /**
     * 类型定义
     */
    UNKNOWN(0),
    DOC(10),
    DOT(11),
    DOCX(20),
    DOCM(21),
    DOTX(22),
    DOTM(23),
    FLAT_OPC(24),
    FLAT_OPC_MACRO_ENABLED(25),
    FLAT_OPC_TEMPLATE(26),
    FLAT_OPC_TEMPLATE_MACRO_ENABLED(27),
    RTF(30),
    WORD_ML(31),
    PDF(40),
    XPS(41),
    XAML_FIXED(42),
    SVG(44),
    HTML_FIXED(45),
    OPEN_XPS(46),
    PS(47),
    PCL(48),
    HTML(50),
    MHTML(51),
    EPUB(52),
    ODT(60),
    OTT(61),
    TEXT(70),
    XAML_FLOW(71),
    XAML_FLOW_PACK(72),
    TIFF(100),
    PNG(101),
    BMP(102),
    EMF(103),
    JPEG(104),
    GIF(105),
    length(35);


    private final int value;

    WordSaveFormat(int value) {
        this.value = value;
    }

    public int getValue() {
        return value;
    }


    public static WordSaveFormat getValueByName(String name) {
        WordSaveFormat[] values = WordSaveFormat.values();
        for (WordSaveFormat wordSaveFormat : values) {
            if (wordSaveFormat.name().equalsIgnoreCase(name)) {
                return wordSaveFormat;
            }
        }
        return UNKNOWN;
    }

}
