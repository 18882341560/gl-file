package com.greelee.gloffice.aspose.model.excel;

import lombok.ToString;

import java.io.Serializable;

import static com.aspose.cells.CellValueType.IS_UNKNOWN;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/6/9
 * @describe:
 */
@ToString
public class ExcelColumnValueUnknown extends AbstractExcelColumn<Object> implements Serializable {
    private static final long serialVersionUID = -4053145702937288457L;
    private Object value;

    public void setValue(Object value) {
        this.value = value;
    }

    /**
     * 类型
     *
     * @return type
     */
    @Override
    public int getType() {
        return IS_UNKNOWN;
    }

    /**
     * 值
     *
     * @return obj
     */
    @Override
    public Object getValue() {
        return this.value;
    }
}
