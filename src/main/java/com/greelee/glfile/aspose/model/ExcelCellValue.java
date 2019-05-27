package com.greelee.glfile.aspose.model;

import com.aspose.cells.Cell;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/5/20
 * @describe:
 */
public interface ExcelCellValue<T> {

    default String getType() {
        return "";
    }

    default T getValue(Cell cell) {
        return null;
    }

}
