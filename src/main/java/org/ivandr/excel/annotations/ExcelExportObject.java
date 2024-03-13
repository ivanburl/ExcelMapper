package org.ivandr.excel.annotations;

import lombok.AllArgsConstructor;
import lombok.NonNull;
import org.dhatim.fastexcel.BorderStyle;
import org.ivandr.excel.enums.ExcelCellHorizontalAlignment;
import org.ivandr.excel.enums.ExcelCellVerticalAlignment;

import java.lang.annotation.*;

@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelExportObject {
    /**
     * Name of header to be generated in Excel export
     */
    String headerName() default "";
    int order();

    @NonNull
    ExcelCellStyle headerStyle() default
    @ExcelCellStyle(isItalic = true, isBold = true, isWrapText = true,
    horizontalAlignment = ExcelCellHorizontalAlignment.CENTER,
    verticalAlignment = ExcelCellVerticalAlignment.CENTER,
    borderStyle = BorderStyle.THICK, fontSize = 16);

    @NonNull
    ExcelCellStyle valueStyle() default
            @ExcelCellStyle(isWrapText = true,
            horizontalAlignment = ExcelCellHorizontalAlignment.CENTER,
            verticalAlignment = ExcelCellVerticalAlignment.CENTER,
            borderStyle = BorderStyle.MEDIUM, fontSize = 12);

    /**
     * fall back value in case null getter returning null
     */
    @NonNull
    String valueFallback() default "";

    /**
     * whether to look for {@link ExcelExportObject} annotation deeper
     */
    boolean isRecursive() default true;
}
