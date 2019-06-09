package com.greelee.gloffice.aspose.excel;

import com.aspose.cells.*;
import com.google.common.base.Strings;
import com.google.common.collect.Lists;
import com.greelee.gloffice.aspose.constant.ExcelSaveFormat;
import com.greelee.gloffice.aspose.exception.OfficeException;
import com.greelee.gloffice.aspose.model.excel.*;

import java.io.InputStream;
import java.util.List;
import java.util.Objects;

import static com.greelee.gloffice.aspose.util.OfficeUtil.dateTimeToLocalDateTime;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/3/30
 * @describe: excel转换
 */
public class ExcelHandler {


    private final static String FOND_NAME = "宋体";
    private final static int FOND_SIZE = 16;
    private final static boolean BOLD = true;

    /**
     * 去除水印，破解了的,这个方法不能公共，每种都有一个license
     */
    private boolean isLicense() {
        InputStream inputStream = ExcelHandler.class.getClassLoader().getResourceAsStream("license.xml");
        License license = new License();
        if (Objects.isNull(inputStream)) {
            throw new OfficeException("not found license.xml");
        }
        license.setLicense(inputStream);
        return true;
    }

    public static void main(String[] args) throws Exception {

        ExcelHandler excelHandler = new ExcelHandler();
        if (!excelHandler.isLicense()) {
            return;
        }
        Workbook workbook = new Workbook();
        WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet worksheet = worksheets.get(0);
        worksheet.setName("第一次测试");
        Cells cells = worksheet.getCells();
        int colnum = 5;
        int rownum = 5;
        Cell cell = cells.get(0, 1);
        cell.setValue("总标题");
        cell.setStyle(excelHandler.createHeaderStyle(workbook));
        cells.merge(0, 0, 1, 5);
        for (int i = 0; i < colnum; i++) {
            cells.get(1, i).setValue("字段标题" + (i + 1));
            cells.get(1, i).setStyle(excelHandler.createHeaderStyle(workbook));
            cells.setRowHeight(0, 25);
            cells.setColumnWidth(i, 50);
        }

        for (int i = 0; i < rownum; i++) {
            for (int j = 0; j < rownum; j++) {
                cells.get(i + 2, j).setValue("哈哈哈哈哈哈哈哈哈哈哈哈哈哈哈哈哈哈哈哈哈哈哈哈哈哈哈");
                cells.get(i + 2, j).setStyle(excelHandler.createHeaderStyle(workbook));
                cells.setRowHeight(i + 2, 20);
                cells.setColumnWidth(i + 2, 38);
            }
        }
        workbook.save("C:\\Users\\gelin\\Desktop\\a.xlsx", SaveFormat.XLSX);
    }


    public void exportExcel(String outPath, int outSheet, String sheetName) {
        Workbook workbook = new Workbook();
        WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet worksheet = worksheets.get(outSheet);
        worksheet.setName(sheetName);
        Cells cells = worksheet.getCells();
    }

    private void createTitltStyle(Cell cell) {

    }

    private Style createHeaderStyle(Workbook workbook) {
        Style style = workbook.createStyle();
        //居中
        style.setHorizontalAlignment(TextAlignmentType.CENTER);
        //自动换行
        style.setTextWrapped(true);
        //单元格 实线
        style.setPattern(BackgroundType.NONE);
        //背景色
//        style.setForegroundArgbColor();
        this.createFont(style);
        return style;
    }

    private void createFont(Style style) {
        Font font = style.getFont();
        font.setBold(BOLD);
        font.setSize(FOND_SIZE);
        font.setName(FOND_NAME);
    }


    /**
     * 得到excel每一列的值
     */
    private AbstractExcelColumn getExcelColumnByType(Cell cell) {
        int type = cell.getType();
        switch (type) {
            case 0:
                ExcelColumnValueBoolean excelColumnValueBoolean = new ExcelColumnValueBoolean();
                excelColumnValueBoolean.setValue(cell.getBoolValue());
                return excelColumnValueBoolean;
            case 1:
                ExcelColumnValueDateTime excelColumnValueDateTime = new ExcelColumnValueDateTime();
                DateTime dateTimeValue = cell.getDateTimeValue();
                excelColumnValueDateTime.setLocalDateTime(dateTimeToLocalDateTime(dateTimeValue));
                return excelColumnValueDateTime;
            case 2:
            case 3:
                return null;
            case 4:
                ExcelColumnValueNumeric excelColumnValueNumeric = new ExcelColumnValueNumeric();
                excelColumnValueNumeric.setNumeric(this.getNumeric(cell));
                return excelColumnValueNumeric;
            case 5:
                ExcelColumnValueString excelColumnValueString = new ExcelColumnValueString();
                excelColumnValueString.setValue(cell.getStringValue());
                return excelColumnValueString;
            default:
                ExcelColumnValueUnknown excelColumnValueUnknown = new ExcelColumnValueUnknown();
                excelColumnValueUnknown.setValue(cell.getValue());
                return excelColumnValueUnknown;
        }
    }

    private Numeric getNumeric(Cell cell) {
        Numeric numeric = new Numeric();
        numeric.setDoubleValue(cell.getDoubleValue());
        numeric.setFloatValue(cell.getFloatValue());
        numeric.setIntValue(cell.getIntValue());
        return numeric;
    }

    /**
     * 获取excel中每一行的值
     */
    public List<ExcelRow> getExcelRowList(String srcPath, int sheet) throws Exception {
        Workbook workbook = new Workbook(srcPath);
        Cells cells = this.getCells(workbook, sheet);
        //有值的行数
        int maxRow = cells.getMaxDataRow() + 1;
        //有值的列数
        int maxColumn = cells.getMaxDataColumn() + 1;
        List<ExcelRow> excelRowList = Lists.newArrayList();
        for (int i = 0; i < maxRow; i++) {
            ExcelRow excelRow = new ExcelRow();
            for (int j = 0; j < maxColumn; j++) {
                Cell cell = cells.get(i, j);
                //得到每一列的值
                AbstractExcelColumn excelColumn = this.getExcelColumnByType(cell);
                //得到该列的公式
                if (Objects.nonNull(excelColumn)) {
                    excelColumn.setFormula(cell.getFormula());
                }
                List<AbstractExcelColumn> excelColumnValueList = excelRow.getExcelColumnValueList();
                if (Objects.isNull(excelColumnValueList)) {
                    List<AbstractExcelColumn> excelColumnList = Lists.newArrayList();
                    excelColumnList.add(excelColumn);
                    excelRow.setExcelColumnValueList(excelColumnList);
                } else {
                    excelRow.getExcelColumnValueList().add(excelColumn);
                }
            }
            excelRowList.add(excelRow);
        }
        return excelRowList;
    }


    /**
     * 获取excel中每一行的值
     */
    public List<List<String>> getExcelContent(String srcPath, int sheet) throws Exception {
        Workbook workbook = new Workbook(srcPath);
        Cells cells = this.getCells(workbook, sheet);
        int maxRow = cells.getMaxDataRow() + 1;
        int maxColumn = cells.getMaxDataColumn() + 1;
        List<List<String>> valueList = Lists.newArrayList();
        for (int i = 0; i < maxRow; i++) {
            List<String> columnList = Lists.newArrayList();
            for (int j = 0; j < maxColumn; j++) {
                columnList.add(cells.get(i, j).getStringValue());
            }
            valueList.add(columnList);
        }
        return valueList;
    }

    private Cells getCells(Workbook workbook, int sheet) {
        WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet worksheet = worksheets.get(sheet);
        return worksheet.getCells();
    }


    /**
     * word转换为其他文件
     *
     * @param inputStream     文件流
     * @param outputPath      存储的地址
     * @param excelSaveFormat 转换的类型
     */
    public void convert(InputStream inputStream, String outputPath, ExcelSaveFormat excelSaveFormat) throws Exception {
        if (Objects.isNull(inputStream) || Strings.isNullOrEmpty(outputPath)) {
            throw new OfficeException("inputStream or outputPath ");
        }
        if (!isLicense()) {
            throw new OfficeException("cells license validation failed");
        }
        Workbook wb = new Workbook(inputStream);
        wb.save(outputPath, excelSaveFormat.getValue());
    }

}
