package com.greelee.gloffice.aspose.model.doc;

import lombok.*;

import java.io.InputStream;
import java.io.Serializable;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/5/12
 * @describe: 模板插入图片, 只支持域的方式
 */
@Getter
@Setter
@ToString
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class WordImage implements Serializable {
    private static final long serialVersionUID = -4585180733987771939L;

    public static final double DEFAULT_WIDTH = 100;
    public static final double DEFAULT_HEIGHT = 100;

    /**
     * 替换的域名/书签名字
     */
    private String key;
    /**
     * 图片流
     */
    private InputStream image;
    /**
     * 插入的图片宽
     */
    private double width;
    /**
     * 插入的图片高
     */
    private double height;

    public double getWidth() {
        return width > 0 ? width : DEFAULT_WIDTH;
    }

    public double getHeight() {
        return height > 0 ? height : DEFAULT_HEIGHT;
    }
}
