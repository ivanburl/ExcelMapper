package org.ivandr;

import lombok.*;
import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;
import org.instancio.Instancio;
import org.instancio.settings.Keys;
import org.instancio.settings.Settings;
import org.ivandr.excel.annotations.ExcelExportObject;
import org.ivandr.excel.mapper.fastexcel.FastExcelMapperFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public class Main {
    @SneakyThrows
    public static void main(String[] args) {
        var start = System.nanoTime();
        var factory = new FastExcelMapperFactory();
        var mapper = factory.createExcelMapperForClass(People.class);
        var end  = System.nanoTime();

        var people = Instancio.of(People.class)
                .withSettings(Settings.defaults().set(Keys.MAX_DEPTH, 10))
                .lenient().create();


        System.out.println((end - start) / 1_000_000);

        var file = new File("./test.xlsx");
        try (OutputStream os = new FileOutputStream(file))
        {
            Workbook wb = new Workbook(os, "MAIN", "0.1");

            Worksheet ws = wb.newWorksheet("example");

            start = System.nanoTime();
            System.out.println(people);
            System.out.println(people.getPeople().size());
            mapper.mapToExcelSheet(ws, 5, 5, people);
            end  = System.nanoTime();
            System.out.println((end - start) / 1_000_000);

            wb.finish();
        }
    }

    @AllArgsConstructor
    @NoArgsConstructor
    @ToString
    @Getter
    @Setter
    public static class People {

        @Getter(onMethod = @__(@ExcelExportObject(order = 0, headerName = "People")))
        private List<Person> people;
    }
}