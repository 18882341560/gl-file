package com.greelee.glfile.aspose.word;

import com.aspose.words.*;
import com.google.common.base.Strings;
import com.greelee.glfile.aspose.constant.WordSaveFormat;
import org.apache.commons.lang3.StringUtils;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Objects;
import java.util.function.Function;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/3/30
 * @describe: word操作类
 */
public class WordOperation {

    private WordOperation() {
    }

    /**
     * 去除水印
     */
    private static boolean isLicense() throws Exception {
        InputStream is = WordOperation.class.getClassLoader().getResourceAsStream("com.aspose.words.lic_2999.xml");
        License license = new License();
        if (Objects.isNull(is)) {
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
        if (!isLicense()) {
            throw new RuntimeException("words license validation failed");
        }
        FileOutputStream os = new FileOutputStream(outputPath);
        Document doc = new Document(inputStream);
        doc.save(os, wordSaveFormat.getValue());
    }

    /**
     * 添加水印文字
     * 代码实例:
     * FileInputStream fileInputStream = new FileInputStream(new File("C:\\Users\\gelin\\Desktop\\基础数据服务.docx"));
     * WordWaterMarkText wordWaterMarkText = WordWaterMarkText.builder()
     * .docInputStream(fileInputStream)
     * .watermarkText("这是水印文字!")
     * .outputPath("C:\\Users\\gelin\\Desktop\\bbbb.docx")
     * .build();
     * WordOperation.addWatermarkText(wordWaterMarkText);
     */
    public static void addWatermarkText(WordWaterMarkText wordWaterMarkText) throws Exception {
        if (paramCheck(wordWaterMarkText)) {
            FileOutputStream os = new FileOutputStream(wordWaterMarkText.getOutputPath());
            Document doc = new Document(wordWaterMarkText.getDocInputStream());
            int top = wordWaterMarkText.getMarginTop() > 0 ? wordWaterMarkText.getMarginTop() : WordWaterMarkText.DEFAULT_MARGIN_TOP;
            int left = wordWaterMarkText.getMarginLeft() > 0 ? wordWaterMarkText.getMarginLeft() : WordWaterMarkText.DEFAULT_MARGIN_LEFT;
            int num = wordWaterMarkText.getNum() > 0 ? wordWaterMarkText.getNum() : WordWaterMarkText.DEFAULT_NUM;
            int loadTopFactor = wordWaterMarkText.getLoadTopFactor() > 0 ? wordWaterMarkText.getLoadTopFactor() : WordWaterMarkText.DEFAULT_LOAD_TOP_FACTOR;
            int loadLeftFactor = wordWaterMarkText.getLoadLeftFactor() > 0 ? wordWaterMarkText.getLoadLeftFactor() : WordWaterMarkText.DEFAULT_LOAD_LEFT_FACTOR;

            for (int i = 0; i < num; i++) {
                int var = left;
                for (int j = 0; j < num; j++) {
                    insertWatermarkText(doc, wordWaterMarkText.getWatermarkText(), top, var, wordWaterMarkText.getFontFamily(),
                            wordWaterMarkText.getWidth(), wordWaterMarkText.getHeight(), wordWaterMarkText.getRotation(),
                            wordWaterMarkText.getColor(), wordWaterMarkText.getRadix());
                    var = var + loadLeftFactor;
                }
                top = top + loadTopFactor;
            }
            String suffix = getSuffix(wordWaterMarkText.getOutputPath());
            doc.save(os, WordSaveFormat.getValueByName(suffix).getValue());
        }
    }


    private static void insertWatermarkText(Document doc, String watermarkText, int top, int left, String fontFamily,
                                            int width, int height, double rotation, String color, int radix) throws Exception {
        insertWatermarkText(doc, watermarkText, fontFamily, width, height, rotation, color, radix, watermark -> {
            // Place the watermark in the page center.
            watermark.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
            watermark.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
            // TOP_BOTTOM : 将所设置位置的内容往上下顶出去
            watermark.setWrapType(WrapType.NONE);
            //垂直对齐
            watermark.setVerticalAlignment(VerticalAlignment.NONE);
            //水平对齐
            watermark.setHorizontalAlignment(HorizontalAlignment.NONE);
            watermark.setTop(top);
            watermark.setLeft(left);
            return null;
        });
    }

    private static void insertWatermarkText(Document doc, String watermarkText, String fontFamily,
                                            int width, int height, double rotation, String color, int radix,
                                            Function<Shape, Object> watermarkPositionConfigFunc) throws Exception {
        // Create a watermark shape. This will be a WordArt shape.
        // You are free to try other shape types as watermarks.
        Shape watermark = new Shape(doc, ShapeType.TEXT_PLAIN_TEXT);
        // Set up the text of the watermark.
        watermark.getTextPath().setText(watermarkText);
        // Set up the text of the watermark.
        // 这里设置为宋体可以保证在转换为PDF时中文不是乱码.
        // default Arial;
        if (StringUtils.isNotBlank(fontFamily)) {
            watermark.getTextPath().setFontFamily(fontFamily);
        } else {
            watermark.getTextPath().setFontFamily(WordWaterMarkText.DEFAULT_FONT_FAMILY);
        }
        // 水印大小
        if (width > 0) {
            watermark.setWidth(width);
        } else {
            watermark.setWidth(WordWaterMarkText.DEFAULT_WIDTH);
        }
        if (height > 0) {
            watermark.setHeight(height);
        } else {
            watermark.setHeight(WordWaterMarkText.DEFAULT_HEIGHT);
        }
        // Text will be directed from the bottom-left to the top-right corner.
        // 左下到右上,角度
        if (rotation > 0) {
            watermark.setRotation(rotation);
        } else {
            watermark.setRotation(WordWaterMarkText.DEFAULT_ROTATION);
        }
        // Remove the following two lines if you need a solid black text.
        String colorStr = WordWaterMarkText.DEFAULT_COLOR;
        if (StringUtils.isNotBlank(color)) {
            colorStr = color;
        }
        int rad = WordWaterMarkText.DEFAULT_RADIX;
        if (radix > 0) {
            rad = radix;
        }
        // Try Color.lightGray to get more Word-style watermark
        watermark.getFill().setColor(new java.awt.Color(Integer.parseInt(colorStr, rad)));
        // Try Color.lightGray to get more Word-style watermark
        watermark.setStrokeColor(new java.awt.Color(Integer.parseInt(colorStr, rad)));
        // Place the watermark in the special location of page .
        watermarkPositionConfigFunc.apply(watermark);
        // Create a new paragraph and append the watermark to this paragraph.
        Paragraph watermarkPara = new Paragraph(doc);
        watermarkPara.appendChild(watermark);
        // Insert the watermark into all headers of each document section.
        for (Section sect : doc.getSections()) {
            // There could be up to three different headers in each section, since we want
            // the watermark to appear on all pages, insert into all headers.
            insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_PRIMARY);
            insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_FIRST);
            insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_EVEN);
        }
        watermark.setZOrder(-100);
    }

    private static void insertWatermarkIntoHeader(Paragraph watermarkPara, Section sect,
                                                  int headerType) {
        HeaderFooter header = sect.getHeadersFooters().getByHeaderFooterType(headerType);
        if (Objects.isNull(header)) {
            // There is no header of the specified type in the current section, create it.
            header = new HeaderFooter(sect.getDocument(), headerType);
            sect.getHeadersFooters().add(header);
        }
        // Insert a clone of the watermark into the header.
        header.appendChild(watermarkPara.deepClone(true));
    }

    private static boolean paramCheck(WordWaterMarkText wordWaterMarkText) {
        if (Objects.nonNull(wordWaterMarkText) && Objects.nonNull(wordWaterMarkText.getDocInputStream())
                && StringUtils.isNotBlank(wordWaterMarkText.getOutputPath()) && StringUtils.isNotBlank(wordWaterMarkText.getWatermarkText())) {
            return true;
        }
        throw new NullPointerException("wordWaterMarkText:" + wordWaterMarkText);
    }


    private static String getSuffix(String path) {
        return path.substring(path.lastIndexOf(".") + 1);
    }

}
