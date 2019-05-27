package com.greelee.glfile.aspose.excel;

import com.aspose.cells.*;
import com.google.common.base.Strings;
import com.greelee.glfile.aspose.constant.ExcelSaveFormat;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/3/30
 * @describe: excel转换
 */
public class ExcelOperation {

    private ExcelOperation() {
    }

    /**
     * 去除水印，破解了的,这个方法不能公共，每种都有一个license
     */
    private static boolean isLicense() {
        InputStream inputStream = ExcelOperation.class.getClassLoader().getResourceAsStream("license.xml");
        License license = new License();
        if (Objects.isNull(inputStream)) {
            throw new NullPointerException("not found license.xml");
        }
        license.setLicense(inputStream);
        return true;
    }

    public static void main(String[] args) throws Exception {
        Workbook workbook = new Workbook("C:\\Users\\gelin\\Desktop\\test.xlsx");
        WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet worksheet = worksheets.get(0);
        Cells cells = worksheet.getCells();
        int maxRow = cells.getMaxDataRow() + 1;
        int maxColumn = cells.getMaxDataColumn() + 1;
        for (int i = 0; i < maxRow; i++) {
            for (int j = 0; j < maxColumn; j++) {
                System.out.println(cells.get(i, j).getType());
            }
        }
    }

    public static String[][] getExcelContent(String srcPath, int sheet) throws Exception {
        Workbook workbook = new Workbook(srcPath);
        WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet worksheet = worksheets.get(sheet);
        Cells cells = worksheet.getCells();
        int maxRow = cells.getMaxDataRow() + 1;
        int maxColumn = cells.getMaxDataColumn() + 1;
        return null;
    }




    /**
     * word转换为其他文件
     *
     * @param inputStream     文件流
     * @param outputPath      存储的地址
     * @param excelSaveFormat 转换的类型
     */
    public static void convert(InputStream inputStream, String outputPath, ExcelSaveFormat excelSaveFormat) throws Exception {
        if (Objects.isNull(inputStream) || Strings.isNullOrEmpty(outputPath)) {
            throw new RuntimeException("inputStream or outputPath ");
        }
        if (!isLicense()) {
            throw new IllegalArgumentException("cells license validation failed");
        }
        Workbook wb = new Workbook(inputStream);
        FileOutputStream fileOS = new FileOutputStream(outputPath);
        wb.save(fileOS, excelSaveFormat.getValue());
    }

}
