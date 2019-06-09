package com.greelee.gloffice.aspose.model.excel;

import lombok.ToString;

import java.io.Serializable;

import static com.aspose.cells.CellValueType.IS_BOOL;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/6/9
 * @describe:
 */
@ToString
public class ExcelColumnValueBoolean extends AbstractExcelColumn<Boolean> implements Serializable {
    private static final long serialVersionUID = -6250127138702973756L;

    private boolean value = false;

    public void setValue(boolean value) {
        this.value = value;
    }

    @Override
    public int getType() {
        return IS_BOOL;
    }

    @Override
    public Boolean getValue() {
        return this.value;
    }
}
