package org.ivandr;

import lombok.*;
import org.ivandr.excel.annotations.ExcelExportObject;

import java.time.LocalDate;
import java.time.temporal.ChronoUnit;
import java.util.Date;
import java.util.List;


@AllArgsConstructor
@Getter
@Setter
@ToString
public class Person {
    @Getter(onMethod = @__(@ExcelExportObject(order = 0, headerName = "First Name")))
    private String name;
    @Getter(onMethod = @__(@ExcelExportObject(order = 1, headerName = "Last name")))
    private String surname;
    @Getter(onMethod = @__(@ExcelExportObject(order = 2, headerName = "Date of birth")))
    private LocalDate birthDate;

    @Getter(onMethod = @__(@ExcelExportObject(order = 3, headerName = "Relatives")))
    private List<Relative> relatives;
    @Getter(onMethod = @__(@ExcelExportObject(order = 5, headerName = "Jobs")))
    private List<Job> jobs;

    @ExcelExportObject(order = 4, headerName = "Age")
    public long getAge() {
        return ChronoUnit.YEARS.between(
                LocalDate.now(),
                birthDate == null ? LocalDate.now() : birthDate
        );
    }
}
