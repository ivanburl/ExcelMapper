package org.ivandr.excel.annotations;

import org.dhatim.fastexcel.BorderStyle;
import org.ivandr.excel.enums.ExcelCellHorizontalAlignment;
import org.ivandr.excel.enums.ExcelCellVerticalAlignment;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCellStyle {
    boolean isBold() default false;

    boolean isItalic() default false;

    boolean isUnderlined() default false;

    boolean isWrapText() default false;

    int fontSize() default 14;

    BorderStyle borderStyle() default BorderStyle.NONE;

    ExcelCellHorizontalAlignment horizontalAlignment() default ExcelCellHorizontalAlignment.GENERAL;

    ExcelCellVerticalAlignment verticalAlignment() default ExcelCellVerticalAlignment.CENTER;

    String cellFormat() default "";
}
