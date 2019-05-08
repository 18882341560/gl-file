package com.greelee.glfile.aspose.excel;

import com.aspose.cells.License;
import com.aspose.cells.Workbook;
import com.google.common.base.Strings;
import com.greelee.glfile.aspose.constant.ExcelSaveFormat;

import java.io.FileOutputStream;
import java.io.InputStream;
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
     * @return
     */
    private static boolean isLicense() {
        InputStream inputStream = ExcelOperation.class.getClassLoader().getResourceAsStream("license.xml");
        License license = new License();
        if(Objects.isNull(inputStream)){
            throw new RuntimeException("not found license.xml");
        }
        license.setLicense(inputStream);
        return true;
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
            throw new RuntimeException("inputStream is null or outputPath is null or empty");
        }
        if(!isLicense()){
            throw new RuntimeException("cells license validation failed");
        }
        Workbook wb = new Workbook(inputStream);
        FileOutputStream fileOS = new FileOutputStream(outputPath);
        wb.save(fileOS, excelSaveFormat.getValue());
    }

}
