package com.greelee.gloffice.aspose.model.doc;

import lombok.*;

import java.io.InputStream;
import java.io.Serializable;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/5/7
 * @describe: word 水印
 */
@Getter
@Setter
@ToString
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class WordWaterMark implements Serializable {

    private static final long serialVersionUID = -6172916064775222514L;


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
     * 水印图片
     */
    private InputStream watermarkImage;
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
     * 字体/图片的长度
     */
    private int width;
    /**
     * 字体/图片的高度
     */
    private int height;
    /**
     * 左下到右上倾斜,角度
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


}
