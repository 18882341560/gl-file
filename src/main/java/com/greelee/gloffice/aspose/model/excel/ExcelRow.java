package com.greelee.gloffice.aspose.model.excel;

import lombok.*;

import java.io.Serializable;
import java.util.List;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/5/20
 * @describe:
 */
@Getter
@Setter
@ToString
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExcelRow implements Serializable {
    private static final long serialVersionUID = 3920614025023646928L;

    /**
     * 字段值
     */
    private List<AbstractExcelColumn> excelColumnValueList;

}
