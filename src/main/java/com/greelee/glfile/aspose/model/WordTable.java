package com.greelee.glfile.aspose.model;

import lombok.*;

import java.io.Serializable;
import java.util.List;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/5/16
 * @describe: word里面的table
 */
@Getter
@Setter
@ToString
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class WordTable implements Serializable {
    private static final long serialVersionUID = -3286384577833557560L;

    /**
     * 这张表的行的集合
     */
    private List<Rows> tableList;

    @Data
    public static class Rows {
        /**
         * 第几行
         */
        private Integer rowIndex;
        /**
         * 这行的列的值
         */
        private List<String> cellList;
    }
}
