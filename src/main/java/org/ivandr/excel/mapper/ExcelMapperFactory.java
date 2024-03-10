package org.ivandr.excel.mapper;

public interface ExcelMapperFactory<WS> {
    <T> ExcelMapper<WS, T> createExcelMapperForClass(Class<T> clazz);
}
