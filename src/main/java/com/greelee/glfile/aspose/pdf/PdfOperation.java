package com.greelee.glfile.aspose.pdf;

import com.aspose.pdf.*;

import java.io.InputStream;
import java.util.Objects;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/5/14
 * @describe: pdf操作类
 */
public class PdfOperation {

    private PdfOperation() {
    }

    /**
     * 去除水印
     */
    private static boolean isLicense() throws Exception {
        InputStream is = Thread.currentThread().getContextClassLoader().getResourceAsStream("license.xml");
        License license = new License();
        if (Objects.isNull(is)) {
            throw new NullPointerException("not found license.xml");
        }
        license.setLicense(is);
        return true;
    }


    public static void main(String[] args) throws Exception {
        if (!isLicense()) {
            System.out.println("666");
        }
        Document document = new Document("C:\\Users\\gelin\\Desktop\\Java8 新特性.pdf");
        document.save("C:\\Users\\gelin\\Desktop\\Java8 新特性.docx", SaveFormat.DocX);
    }

}
