package com.greelee.glfile.aspose.word;

import com.aspose.words.*;
import com.google.common.base.Strings;
import com.google.common.collect.Lists;
import com.greelee.glfile.aspose.constant.WordSaveFormat;
import com.greelee.glfile.aspose.model.WordImage;
import com.greelee.glfile.aspose.model.WordReplace;
import com.greelee.glfile.aspose.model.WordWaterMark;
import com.greelee.glfile.aspose.util.DocumentUtil;
import org.apache.commons.lang3.StringUtils;

import java.io.*;
import java.time.Instant;
import java.util.List;
import java.util.Objects;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/3/30
 * @describe: word操作类
 */
public class WordOperation {


    private static final String DEFAULT_FONT_FAMILY = "宋体";
    private static final String DEFAULT_COLOR = "9FB6CD";
    private static final int DEFAULT_RADIX = 16;
    private static final int DEFAULT_WIDTH = 100;
    private static final int DEFAULT_HEIGHT = 100;
    /**
     * 左下到右上,角度
     */
    private static final double DEFAULT_ROTATION = 0;
    private static final int DEFAULT_MARGIN_TOP = 250;
    private static final int DEFAULT_MARGIN_LEFT = 450;

    private static final String DEFAULT_SUFFIX_NAME = "png";

    private WordOperation() {
    }

    /**
     * 去除水印
     */
    private static boolean isLicense() throws Exception {
        InputStream is = WordOperation.class.getClassLoader().getResourceAsStream("com.aspose.words.lic_2999.xml");
        License license = new License();
        if (Objects.isNull(is)) {
            throw new NullPointerException("not found license.xml");
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
     * 将word的图片获取出来,保存
     *
     * @param docInputStream word文件流
     * @param directory      要保存的目录
     * @param suffixName     后缀名称
     * @return 文件路径名称
     */
    public static List<String> saveImagesToDocument(InputStream docInputStream, String directory, String suffixName) throws Exception {
        if (Objects.nonNull(directory) && StringUtils.isNotBlank(directory)) {
            if (!StringUtils.isNotBlank(suffixName)) {
                suffixName = DEFAULT_SUFFIX_NAME;
            }
            Document doc = new Document(docInputStream);
            closeStream(docInputStream);
            return saveImage(doc, directory, suffixName);
        }
        return null;
    }

    /**
     * 将word的图片获取出来,保存
     *
     * @param docFileName word路径
     * @param directory   要保存的目录
     * @param suffixName  后缀名称
     * @return 文件路径名称
     */
    public static List<String> saveImagesToDocument(String docFileName, String directory, String suffixName) throws Exception {
        if (StringUtils.isNotBlank(docFileName) && StringUtils.isNotBlank(directory)) {
            if (!StringUtils.isNotBlank(suffixName)) {
                suffixName = DEFAULT_SUFFIX_NAME;
            }
            Document doc = new Document(docFileName);
            return saveImage(doc, directory, suffixName);
        }
        return null;
    }


    /**
     * 获取word里面的所有图片
     */
    public static List<byte[]> getDocumentImages(InputStream docInputStream) throws Exception {
        if (Objects.nonNull(docInputStream)) {
            Document doc = new Document(docInputStream);
            closeStream(docInputStream);
            return getImageBytes(doc);
        }
        return null;
    }

    /**
     * 获取word里面的所有图片
     */
    public static List<byte[]> getDocumentImages(String fileName) throws Exception {
        if (StringUtils.isNotBlank(fileName)) {
            Document doc = new Document(fileName);
            return getImageBytes(doc);
        }
        return null;
    }


    private static List<String> saveImage(Document document, String directory, String suffixName) throws Exception {
        NodeCollection childNodes = document.getChildNodes(NodeType.SHAPE, true);
        Node[] nodes = childNodes.toArray();
        List<String> list = Lists.newArrayList();
        if (Objects.nonNull(nodes) && nodes.length > 0) {
            for (Node node : nodes) {
                Shape shape = (Shape) node;
                ImageData imageData = shape.getImageData();
                File file = new File(directory);
                if (!file.exists()) {
                    if (!file.mkdirs()) {
                        throw new NullPointerException(directory);
                    }
                }
                String fileName = directory + File.separator + Instant.now().getEpochSecond() + "." + suffixName;
                imageData.save(fileName);
                list.add(fileName);
            }
            return list;
        }
        return null;
    }

    private static List<byte[]> getImageBytes(Document doc) throws Exception {
        NodeCollection childNodes = doc.getChildNodes(NodeType.SHAPE, true);
        Node[] nodes = childNodes.toArray();
        List<byte[]> list = Lists.newArrayList();
        if (Objects.nonNull(nodes) && nodes.length > 0) {
            for (Node node : nodes) {
                Shape shape = (Shape) node;
                ImageData imageData = shape.getImageData();
                list.add(imageData.getImageBytes());
            }
            return list;
        }
        return null;
    }


    /**
     * 自定义域的生成模板
     */
    public static void generateTemplateToMailMerge(WordReplace wordReplace) throws Exception {
        if (generateTemplateParamCheck(wordReplace)) {
            List filedNames = Objects.requireNonNull(DocumentUtil.getMapKeyValueList(wordReplace.getReplaceTextData())).getKList();
            List values = Objects.requireNonNull(DocumentUtil.getMapKeyValueList(wordReplace.getReplaceTextData())).getVList();
            Document doc = new Document(wordReplace.getDocInputStream());
            doc.getMailMerge().execute((String[]) filedNames.toArray(new String[0]), values.toArray());
            insertImage(doc, wordReplace.getWordImageList());
            saveDocument(doc, wordReplace.getOutputPath());
            closeStream(wordReplace.getDocInputStream());
        }
    }

    private static void insertImage(Document doc, List<WordImage> wordImageList) throws Exception {
        if (DocumentUtil.isNotEmpty(wordImageList)) {
            DocumentBuilder builder = new DocumentBuilder(doc);
            for (WordImage wordImage : wordImageList) {
                if (Objects.nonNull(wordImage.getImage()) && StringUtils.isNotBlank(wordImage.getKey())) {
                    builder.moveToMergeField(wordImage.getKey());
                    builder.insertImage(wordImage.getImage(), wordImage.getWidth(), wordImage.getHeight());
                    closeStream(wordImage.getImage());
                }
            }
        }
    }


    /**
     * 自定义替换符的生成模板
     */
    public static void generateTemplate(WordReplace wordReplace) throws Exception {
        if (generateTemplateParamCheck(wordReplace)) {
            Document doc = new Document(wordReplace.getDocInputStream());
            FindReplaceOptions options = new FindReplaceOptions();
            Range range = doc.getRange();
            wordReplace.getReplaceTextData().forEach((k, v) -> {
                try {
                    range.replace(k, v, options);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            });
            saveDocument(doc, wordReplace.getOutputPath());
            closeStream(wordReplace.getDocInputStream());
        }
    }

    private static boolean generateTemplateParamCheck(WordReplace wordReplace) {
        if (Objects.nonNull(wordReplace) && Objects.nonNull(wordReplace.getDocInputStream())
                && StringUtils.isNotBlank(wordReplace.getOutputPath())
                && DocumentUtil.isNotEmpty(wordReplace.getReplaceTextData())) {
            return true;
        }
        throw new NullPointerException("wordReplace:" + wordReplace);
    }


    /**
     * 添加水印文字,多了会将文字往下挤
     * 代码实例:
     * FileInputStream fileInputStream = new FileInputStream(new File("C:\\Users\\gelin\\Desktop\\基础数据服务.docx"));
     * WordWaterMark wordWaterMark = WordWaterMark.builder()
     * .docInputStream(fileInputStream)
     * .watermarkText("这是水印文字!")
     * .outputPath("C:\\Users\\gelin\\Desktop\\bbbb.docx")
     * .build();
     * WordOperation.addWatermarkText(wordWaterMark);
     */
    public static void addWatermarkText(WordWaterMark wordWaterMark) throws Exception {
        if (paramCheckWatermarkText(wordWaterMark)) {
            Document doc = new Document(wordWaterMark.getDocInputStream());
            initWordWaterMark(wordWaterMark);
            setWatermarkText(doc, wordWaterMark.getWatermarkText(), wordWaterMark.getMarginTop(), wordWaterMark.getMarginLeft(),
                    wordWaterMark.getFontFamily(), wordWaterMark.getWidth(), wordWaterMark.getHeight(),
                    wordWaterMark.getRotation(), wordWaterMark.getColor(), wordWaterMark.getRadix());
            saveDocument(doc, wordWaterMark.getOutputPath());
            close(wordWaterMark);
        }
    }


    /**
     * 添加水印图片,多了会将文字往下挤
     * 代码实例:
     * FileInputStream fileInputStream = new FileInputStream(new File("C:\\Users\\gelin\\Desktop\\基础数据服务.docx"));
     * WordWaterMark wordWaterMark = WordWaterMark.builder()
     * .docInputStream(fileInputStream)
     * .watermarkImage() 图片流
     * .outputPath("C:\\Users\\gelin\\Desktop\\bbbb.docx")
     * .build();
     * WordOperation.addWatermarkText(wordWaterMark);
     */
    public static void addWatermarkImage(WordWaterMark wordWaterMark) throws Exception {
        if (paramCheckWatermarkImage(wordWaterMark)) {
            Document doc = new Document(wordWaterMark.getDocInputStream());
            initWordWaterMark(wordWaterMark);
            Shape watermark = new Shape(doc, ShapeType.IMAGE);
            watermark.setBehindText(true);
            watermark.getImageData().setImage(wordWaterMark.getWatermarkImage());
            insertWatermarkImage(doc, watermark, wordWaterMark.getMarginTop(), wordWaterMark.getMarginLeft(),
                    wordWaterMark.getWidth(), wordWaterMark.getHeight(), wordWaterMark.getRotation());
            saveDocument(doc, wordWaterMark.getOutputPath());
            close(wordWaterMark);
        }
    }

    private static void saveDocument(Document document, String outputPath) throws Exception {
        String suffix = DocumentUtil.getSuffix(outputPath);
        document.save(outputPath, WordSaveFormat.getValueByName(suffix).getValue());
    }


    private static void close(WordWaterMark wordWaterMark) throws IOException {
        closeStream(wordWaterMark.getWatermarkImage());
        closeStream(wordWaterMark.getDocInputStream());
    }

    private static void closeStream(InputStream inputStream) {
        if (Objects.nonNull(inputStream)) {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static void initWordWaterMark(WordWaterMark wordWaterMark) {
        if (wordWaterMark.getMarginTop() <= 0) {
            wordWaterMark.setMarginTop(DEFAULT_MARGIN_TOP);
        }
        if (wordWaterMark.getMarginLeft() <= 0) {
            wordWaterMark.setMarginLeft(DEFAULT_MARGIN_LEFT);
        }
    }


    private static void insertWatermarkImage(Document doc, Shape watermark, int top, int left,
                                             int width, int height, double rotation) throws Exception {
        setShape(watermark, top, left, width, height, rotation);
        setParagraph(doc, watermark);
    }


    private static void setWatermarkText(Document doc, String watermarkText, int top, int left, String fontFamily,
                                         int width, int height, double rotation, String color, int radix) throws Exception {
        Shape watermark = new Shape(doc, ShapeType.TEXT_PLAIN_TEXT);
        setShape(watermark, top, left, width, height, rotation);
        setWatermarkText(doc, watermarkText, fontFamily, color, radix, watermark);
    }


    private static void setShape(Shape shape, int top, int left, int width, int height, double rotation) throws Exception {
        // Place the watermark in the page center.
        shape.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        shape.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
        // TOP_BOTTOM : 将所设置位置的内容往上下顶出去
        shape.setWrapType(WrapType.NONE);
        //垂直对齐
        shape.setVerticalAlignment(VerticalAlignment.NONE);
        //水平对齐
        shape.setHorizontalAlignment(HorizontalAlignment.NONE);
        // 水印大小
        if (width > 0) {
            shape.setWidth(width);
        } else {
            shape.setWidth(DEFAULT_WIDTH);
        }
        if (height > 0) {
            shape.setHeight(height);
        } else {
            shape.setHeight(DEFAULT_HEIGHT);
        }
        // Text will be directed from the bottom-left to the top-right corner.
        // 左下到右上,角度
        if (rotation > 0) {
            shape.setRotation(rotation);
        } else {
            shape.setRotation(DEFAULT_ROTATION);
        }
        if (top > 0) {
            shape.setTop(top);
        } else {
            shape.setTop(DEFAULT_MARGIN_TOP);
        }

        if (left > 0) {
            shape.setLeft(left);
        } else {
            shape.setLeft(DEFAULT_MARGIN_LEFT);
        }
    }


    private static void setWatermarkText(Document doc, String watermarkText, String fontFamily, String color, int radix,
                                         Shape watermark) throws Exception {
        // Create a watermark shape. This will be a WordArt shape.
        // You are free to try other shape types as watermarks.
        // Set up the text of the watermark.
        watermark.getTextPath().setText(watermarkText);
        // Set up the text of the watermark.
        // 这里设置为宋体可以保证在转换为PDF时中文不是乱码.
        // default Arial;
        if (StringUtils.isNotBlank(fontFamily)) {
            watermark.getTextPath().setFontFamily(fontFamily);
        } else {
            watermark.getTextPath().setFontFamily(DEFAULT_FONT_FAMILY);
        }
        // Remove the following two lines if you need a solid black text.
        String colorStr = DEFAULT_COLOR;
        if (StringUtils.isNotBlank(color)) {
            colorStr = color;
        }
        int rad = DEFAULT_RADIX;
        if (radix > 0) {
            rad = radix;
        }
        // Try Color.lightGray to get more Word-style watermark
        watermark.getFill().setColor(new java.awt.Color(Integer.parseInt(colorStr, rad)));
        // Try Color.lightGray to get more Word-style watermark
        watermark.setStrokeColor(new java.awt.Color(Integer.parseInt(colorStr, rad)));
        // Place the watermark in the special location of page .
        // Create a new paragraph and append the watermark to this paragraph.
        setParagraph(doc, watermark);
    }

    private static void setParagraph(Document doc, Shape watermark) {
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

    private static boolean paramCheckWatermarkText(WordWaterMark wordWaterMark) {
        if (Objects.nonNull(wordWaterMark) && Objects.nonNull(wordWaterMark.getDocInputStream())
                && StringUtils.isNotBlank(wordWaterMark.getOutputPath()) && StringUtils.isNotBlank(wordWaterMark.getWatermarkText())) {
            return true;
        }
        throw new NullPointerException("wordWaterMark:" + wordWaterMark);
    }

    private static boolean paramCheckWatermarkImage(WordWaterMark wordWaterMark) {
        if (Objects.nonNull(wordWaterMark) && Objects.nonNull(wordWaterMark.getDocInputStream())
                && StringUtils.isNotBlank(wordWaterMark.getOutputPath()) && Objects.nonNull(wordWaterMark.getWatermarkImage())) {
            return true;
        }
        throw new NullPointerException("wordWaterMark:" + wordWaterMark);
    }


}
