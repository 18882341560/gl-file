package com.greelee.glfile.aspose.word;

import lombok.*;

import java.io.InputStream;
import java.io.Serializable;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/5/7
 * @describe: word 文字水印
 */
@Getter
@Setter
@ToString
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class WordWaterMarkText implements Serializable {

    private static final long serialVersionUID = -6172916064775222514L;

    public static final String DEFAULT_FONT_FAMILY = "宋体";
    public static final String DEFAULT_COLOR = "9FB6CD";
    public static final int DEFAULT_RADIX = 16;
    public static final int DEFAULT_WIDTH = 100;
    public static final int DEFAULT_HEIGHT = 100;
    /**
     * 左下到右上,角度
     */
    public static final double DEFAULT_ROTATION = 0;
    /**
     * 这一页有多少排水印,在高度范围内
     */
    public static final int DEFAULT_NUM = 3;
    public static final int DEFAULT_MARGIN_TOP = 50;
    public static final int DEFAULT_MARGIN_LEFT = 50;
    public static final int DEFAULT_LOAD_TOP_FACTOR = 150;
    public static final int DEFAULT_LOAD_LEFT_FACTOR = 200;


    /**
     * word文档流
     */
    private InputStream docInputStream;
    /**
     * 输出的路径
     */
    private String outputPath;
    /**
     * 水印文字
     */
    private String watermarkText;
    /**
     * 字体
     */
    private String fontFamily;
    /**
     * 距离顶部的距离
     */
    private int marginTop;
    /**
     * 距离左部的距离
     */
    private int marginLeft;
    /**
     * 字体的长度
     */
    private int width;
    /**
     * 字体的高度
     */
    private int height;
    /**
     * 左下到右上,角度
     */
    private double rotation;
    /**
     * 颜色,不带#
     */
    private String color;
    /**
     * the radix to be used while parsing
     */
    private int radix;
    /**
     * 这一页有多少排水印,在高度范围内
     */
    private int num;
    /**
     * 加载的垂直高度的因子
     */
    private int loadTopFactor;
    /**
     * 加载的水平长度的因子
     */
    private int loadLeftFactor;


}
