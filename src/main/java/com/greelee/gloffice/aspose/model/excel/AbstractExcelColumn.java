package com.greelee.gloffice.aspose.model.excel;


import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import static com.aspose.cells.CellValueType.IS_UNKNOWN;

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
public abstract class AbstractExcelColumn<T> {

    private String formula;

    /**
     * 类型
     *
     * @return type
     */
    protected int getType() {
        return IS_UNKNOWN;
    }

    /**
     * 值
     *
     * @return obj
     */
    protected T getValue() {
        return null;
    }

}
