package org.ivandr.excel.enums;

import lombok.AllArgsConstructor;
import lombok.Getter;

@AllArgsConstructor
@Getter
public enum ExcelCellHorizontalAlignment {
    GENERAL("general"),
    CENTER("center"),
    LEFT("left"),
    RIGHT("right"),
    JUSTIFY("justify"),
    FILL("fill"),
    CENTER_CONTINUOUS("centerContinuous"),
    DISTRIBUTED( "distributed");

    /**
     * <a href="https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc880467(v=office.14)?redirectedfrom=MSDN">Link to names</a>
     */
    private final String horizontalAlignmentTag;
}
