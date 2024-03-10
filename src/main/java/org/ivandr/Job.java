package org.ivandr;

import lombok.*;
import org.ivandr.excel.annotations.ExcelExportObject;

import java.time.LocalDate;
import java.util.List;


@AllArgsConstructor
@Getter
@Setter
@ToString
public class Job {
    @Getter(onMethod = @__(@ExcelExportObject(order = 0, headerName = "Company name")))
    private String company;
    @Getter(onMethod = @__(@ExcelExportObject(order = 1, headerName = "Position Name")))
    private String position;

    @Getter(onMethod = @__(@ExcelExportObject(order = 2, headerName = "Manager")))
    private Manager manager;

    @Getter(onMethod = @__(@ExcelExportObject(order = 3, headerName = "Start date of job")))
    private LocalDate startDate;
    @Getter(onMethod = @__(@ExcelExportObject(order = 4, headerName = "End date of job")))
    private LocalDate endDate;

    @ToString
    public static class Manager extends Person {
        public Manager(String name, String surname, LocalDate birthDate, List<Relative> relatives, List<Job> jobs) {
            super(name, surname, birthDate, relatives, jobs);
        }

        @Override
        public List<Job> getJobs() {
            return super.getJobs();
        }
        @Override
        public List<Relative> getRelatives() {
            return super.getRelatives();
        }
    }
}
