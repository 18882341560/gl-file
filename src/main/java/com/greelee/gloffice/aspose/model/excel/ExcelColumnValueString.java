package com.greelee.gloffice.aspose.model.excel;

import lombok.ToString;

import java.io.Serializable;

import static com.aspose.cells.CellValueType.IS_STRING;


/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/5/20
 * @describe:
 */
@ToString
public class ExcelColumnValueString extends AbstractExcelColumn<String> implements Serializable {
    private static final long serialVersionUID = -1141290195372714928L;

    private String value;

    public void setValue(String value) {
        this.value = value;
    }

    @Override
    public int getType() {
        return IS_STRING;
    }

    @Override
    public String getValue() {
        return this.value;
    }
}
