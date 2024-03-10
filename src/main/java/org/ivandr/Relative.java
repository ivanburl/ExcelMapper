package org.ivandr;

import lombok.*;
import org.ivandr.excel.annotations.ExcelExportObject;

import java.time.LocalDate;
import java.util.List;

@Getter
@Setter
@ToString
public class Relative extends Person {

    public Relative(String name, String surname, LocalDate birthDate, List<Relative> relatives, List<Job> jobs) {
        super(name, surname, birthDate, relatives, jobs);
    }

    @Override
    public List<Relative> getRelatives() {
        return super.getRelatives();
    }

    @Override
    public List<Job> getJobs() {
        return super.getJobs();
    }
}
