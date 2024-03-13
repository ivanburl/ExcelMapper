package org.ivandr.excel.mapper;

import lombok.NonNull;

public interface ExcelMapper<WS, T> {
    /**
     * Creates the table in worksheet, using configuration of object class, and his values
     * @param startRow left upper row of table
     * @param startColumn left upper column of table
     * @param worksheet worksheet to which would dbe done exporting
     * @param object object which would be exported to specific worksheet
     */
    void mapToExcelSheet(
            @NonNull WS worksheet,
            int startRow, int startColumn,
            T object);
}
