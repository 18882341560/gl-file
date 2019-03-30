package com.greelee.glfile.aspose.constant;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/3/30
 * @describe: 类型格式
 */
public enum ExcelSaveFormat {

    /**
     * 类型定义
     */
    CSV(1),
    XLSX(6),
    XLSM(7),
    XLTX(8),
    XLTM(9),
    XLAM(10),
    TSV(11),
    TAB_DELIMITED(11),
    HTML(12),
    M_HTML(17),
    ODS(14),
    EXCEL_97_TO_2003(5),
    SPREADSHEET_ML(15),
    XLSB(16),
    AUTO(0),
    UNKNOWN(256),
    PDF(13),
    XPS(20),
    TIFF(21),
    SVG(22),
    DIF(30),
    NUMBERS(56);


    private final int value;

    ExcelSaveFormat(int value) {
        this.value = value;
    }

    public int getValue() {
        return value;
    }

}
