package org.ivandr.excel.mapper;

import lombok.NonNull;

public record ExcelCellCoordinates(int row, int column) {
    public static ExcelCellCoordinates sum(@NonNull ExcelCellCoordinates a, @NonNull ExcelCellCoordinates b) {
        return new ExcelCellCoordinates(a.row + b.row, a.column + b.column);
    }
}
