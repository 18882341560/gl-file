package com.greelee.glfile.aspose.model;

import lombok.*;

import java.io.InputStream;
import java.io.Serializable;
import java.util.List;
import java.util.Map;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/5/12
 * @describe: word 模板生成
 */
@Getter
@Setter
@ToString
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class WordReplace implements Serializable {
    private static final long serialVersionUID = 451166379435746754L;

    /**
     * word文档流
     */
    private InputStream docInputStream;
    /**
     * 输出的路径
     */
    private String outputPath;
    /**
     * 纯文本替换,使用域替换,key只需要替换的名称 例子:name/使用自定义的需要整个替换的字符 例子: ${name}
     */
    private Map<String, String> replaceTextData;
    /**
     * 要替换成图片,只支持域的方式
     */
    private List<WordImage> wordImageList;
}
