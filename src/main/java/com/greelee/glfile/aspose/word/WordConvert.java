package com.greelee.glfile.aspose.word;

import com.aspose.words.Document;
import com.aspose.words.License;
import com.google.common.base.Strings;
import com.greelee.glfile.aspose.constant.WordSaveFormat;

import java.io.*;
import java.util.Objects;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/3/30
 * @describe: word转换
 */
public class WordConvert {

    private WordConvert() {
    }

    /**
     * 去除水印，
     * @return
     * @throws Exception
     */
    private static boolean isLicense() throws Exception {
        InputStream is = WordConvert.class.getClassLoader().getResourceAsStream("com.aspose.words.lic_2999.xml");
        License license = new License();
        if(Objects.isNull(is)){
            throw new RuntimeException("not found license.xml");
        }
        license.setLicense(is);
        return true;
    }


    /**
     * word转换为其他文件
     *
     * @param inputStream    文件流
     * @param outputPath     存储的地址
     * @param wordSaveFormat 转换的类型
     */
    public static void convert(InputStream inputStream, String outputPath, WordSaveFormat wordSaveFormat) throws Exception {
        if (Objects.isNull(inputStream) || Strings.isNullOrEmpty(outputPath)) {
            throw new RuntimeException("inputStream is null or outputPath is null or empty");
        }
        if(!isLicense()){
            throw new RuntimeException("words license validation failed");
        }
        FileOutputStream os = new FileOutputStream(outputPath);
        Document doc = new Document(inputStream);
        doc.save(os, wordSaveFormat.getValue());
    }

}
