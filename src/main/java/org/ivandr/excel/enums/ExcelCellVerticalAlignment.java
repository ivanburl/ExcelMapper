package org.ivandr.excel.enums;

import lombok.AllArgsConstructor;
import lombok.Getter;

@AllArgsConstructor
@Getter
public enum ExcelCellVerticalAlignment {

    TOP("top"),
    CENTER("center"),
    BOTTOM( "bottom"),
    JUSTIFY( "justify"),
    DISTRIBUTED("distributed");

    /**
     * <a href="https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc802119(v=office.14)?redirectedfrom=MSDN">...</a>
     */
    private final String verticalAlignmentFastExcelTag;
}
