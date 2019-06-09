package com.greelee.gloffice.aspose.model.excel;

import lombok.ToString;

import java.io.Serializable;
import java.time.LocalDateTime;

import static com.aspose.cells.CellValueType.IS_DATE_TIME;


/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/5/20
 * @describe:
 */
@ToString
public class ExcelColumnValueDateTime extends AbstractExcelColumn<LocalDateTime> implements Serializable {
    private static final long serialVersionUID = -1141290195372714928L;

    private LocalDateTime localDateTime;

    public void setLocalDateTime(LocalDateTime localDateTime) {
        this.localDateTime = localDateTime;
    }

    @Override
    public int getType() {
        return IS_DATE_TIME;
    }

    @Override
    public LocalDateTime getValue() {
        return this.localDateTime;
    }
}
