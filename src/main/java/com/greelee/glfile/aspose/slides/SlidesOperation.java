package com.greelee.glfile.aspose.slides;

import com.aspose.slides.*;
import com.google.common.base.Strings;
import com.greelee.glfile.aspose.constant.SlidesSaveFormat;

import java.io.*;
import java.util.Objects;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/3/30
 * @describe: ppt转换
 */
public class SlidesOperation {

    private SlidesOperation() {
    }

    /**
     * 去除水映 破解了的
     */
    private static boolean isLicense() {
        InputStream licenseIs = SlidesOperation.class.getClassLoader().getResourceAsStream("license.xml");
        License license = new License();
        if (Objects.isNull(licenseIs)) {
            throw new NullPointerException("not found license.xml");
        }
        license.setLicense(licenseIs);
        return true;
    }

    /**
     * word转换为其他文件
     *
     * @param inputStream      文件流
     * @param outputPath       存储的地址
     * @param slidesSaveFormat 转换的类型
     */
    public static void convert(InputStream inputStream, String outputPath, SlidesSaveFormat slidesSaveFormat) throws Exception {
        if (Objects.isNull(inputStream) || Strings.isNullOrEmpty(outputPath)) {
            throw new RuntimeException("inputStream is null or outputPath is null or empty");
        }
        if (!isLicense()) {
            throw new RuntimeException("slides license validation failed");
        }
        Presentation pres = new Presentation(inputStream);
        FileOutputStream fileOS = new FileOutputStream(outputPath);
        pres.save(fileOS, slidesSaveFormat.getValue());
    }

}
