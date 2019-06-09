package com.greelee.gloffice.aspose.model.excel;

import lombok.ToString;

import java.io.Serializable;

import static com.aspose.cells.CellValueType.IS_NUMERIC;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/6/9
 * @describe:
 */
@ToString
public class ExcelColumnValueNumeric extends AbstractExcelColumn<Numeric> implements Serializable {
    private static final long serialVersionUID = 4495373511442664210L;

    private Numeric numeric;

    public void setNumeric(Numeric numeric) {
        this.numeric = numeric;
    }

    /**
     * 值
     */
    @Override
    public Numeric getValue() {
        return this.numeric;
    }

    /**
     * 类型
     */
    @Override
    public int getType() {
        return IS_NUMERIC;
    }
}
