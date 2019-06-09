package com.greelee.gloffice.aspose.model.doc;

import com.google.common.collect.Lists;
import lombok.*;

import java.io.Serializable;
import java.util.List;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/5/16
 * @describe: 页眉页脚
 */
@Getter
@Setter
@ToString
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class WordHeaderFooter implements Serializable {
    private static final long serialVersionUID = -6366849319691849034L;

    private static List<String> list = Lists.newArrayList();

    private String header;
    private String footer;

    /**
     * 可能有问题的页眉页脚没有解析出来就把所有的全部放到这里
     */
    List<String> errorList;

    public List<String> getErrorList() {
        return list;
    }

    public void setErrorList(List<String> errorList) {
        list = errorList;
    }
}
