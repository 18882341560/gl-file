package com.greelee.glfile.aspose.constant;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/3/30
 * @describe: 幻灯片类型格式
 */
public enum SlidesSaveFormat {

    /**
     * 类型定义
     */
    PPT(0),
    PPS(0),
    PDF(1),
    XPS(2),
    PPTX(3),
    PPSX(4),
    TIFF(5),
    ODP(6),
    PPTM(7),
    PPSM(9),
    POTX(10),
    POTM(11),
    PDFNOTES(12),
    HTML(13),
    TIFFNOTES(14);


    private final int value;

    SlidesSaveFormat(int value) {
        this.value = value;
    }

    public int getValue() {
        return value;
    }

}
