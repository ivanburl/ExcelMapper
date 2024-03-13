package org.ivandr.excel.mapper;

import lombok.*;
import org.checkerframework.checker.units.qual.N;
import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;
import org.instancio.Instancio;
import org.instancio.settings.Keys;
import org.instancio.settings.Settings;
import org.ivandr.Main;
import org.ivandr.Person;
import org.ivandr.excel.annotations.ExcelExportObject;
import org.ivandr.excel.mapper.fastexcel.FastExcelMapperFactory;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.RepeatedTest;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.List;
import java.util.Objects;

import static org.junit.jupiter.api.Assertions.*;

class ExcelMapperTest {

    @Test
    @SneakyThrows
    void mapNullToExcelSheet() {
        try (var wb = createAndWriteWorkbook(SimpleClass.class, null,
                "simple_object"
        )) {
            Assertions.assertNotNull(wb);
        }
        //TODO compare it with some result
    }

    @SneakyThrows
    @RepeatedTest(10)
    void mapSimpleObjectToExcelSheet() {
        var object = Instancio.of(SimpleClass.class)
                .withSettings(Settings.defaults())
                .lenient().create();

        try (var wb = createAndWriteWorkbook(SimpleClass.class, object,
                "simple_object"
        )) {
            Assertions.assertNotNull(wb);
        }
        //TODO compare it with some result
    }

    @SneakyThrows
    @RepeatedTest(10)
    void mapObjectWithListToExcelSheet() {
        var object = Instancio.of(SimpleClassWithList.class)
                .withSettings(Settings.defaults()
                        .set(Keys.COLLECTION_MIN_SIZE, 5)
                        .set(Keys.COLLECTION_MAX_SIZE, 10))
                .lenient().create();
        System.out.println(object);

        try (var wb = createAndWriteWorkbook(SimpleClassWithList.class, object,
                "simple_object_with_list"
        )) {

            Assertions.assertNotNull(wb);
        }
    }

    @SneakyThrows
    @RepeatedTest(10)
    void mapSimpleListOfObjectToSheet() {
        var object = Instancio.of(SimpleClassListWrapper.class)
                .withSettings(Settings.defaults()
                        .set(Keys.COLLECTION_MIN_SIZE, 10)
                        .set(Keys.COLLECTION_MAX_SIZE, 30))
                .lenient().create();
        System.out.println(object);

        try (var wb = createAndWriteWorkbook(SimpleClassListWrapper.class, object,
                "list_of_simple_class"
        )) {

            Assertions.assertNotNull(wb);
        }
    }

    @SneakyThrows
    @RepeatedTest(10)
    void mapObjectWithSimpleRecursionSheet() {
        var object = Instancio.of(Person.class)
                .withSettings(Settings.defaults()
                        .set(Keys.MAX_DEPTH, 10)
                        .set(Keys.COLLECTION_MIN_SIZE, 10)
                        .set(Keys.COLLECTION_MAX_SIZE, 30))
                .lenient().create();
        System.out.println(object);


        try (var wb = createAndWriteWorkbook(Person.class, object,
                "recursive_simple"
        )) {

            Assertions.assertNotNull(wb);
        }
    }

    @RepeatedTest(10)
    @SneakyThrows
    void mapComplexObjectToExcelSheet() {
        var people = Instancio.of(Main.People.class)
                .withSettings(Settings.defaults()
                        .set(Keys.MAX_DEPTH, 10)
                        .set(Keys.COLLECTION_MIN_SIZE, 10)
                        .set(Keys.COLLECTION_MAX_SIZE, 30))
                .lenient().create();
        System.out.println(people);


        try (var wb = createAndWriteWorkbook(Main.People.class, people,
                "complex_export"
        )) {

            Assertions.assertNotNull(wb);
        }
    }

    @SneakyThrows
    private <T> Workbook createAndWriteWorkbook(
            Class<T> clazz,
            T value,
            @NonNull String excelFileName
    ) {
        var resource = Paths.get(Objects.requireNonNull(getClass().getClassLoader().getResource(".")).toURI());
        Path filePath = Paths.get("%s/%s.xlsx".formatted(resource.toAbsolutePath(), excelFileName));

        var file = filePath.toFile();

        var factory = new FastExcelMapperFactory();
        var mapper = factory.createExcelMapperForClass(clazz);

        Workbook wb;
        try (OutputStream os = new FileOutputStream(file)) {
            wb = new Workbook(os, getClass().getName(), "0.1");

            Worksheet ws = wb.newWorksheet(clazz.getSimpleName());
            mapper.mapToExcelSheet(ws, 4, 4, value);
            wb.finish();
        }

        return wb;
    }

    public enum SimpleEnum {
        VARIANT_A, VARIANT_B, VARIANT_C
    }

    @AllArgsConstructor
    @Getter
    @ToString(callSuper = true)
    public static class SimpleClass {
        @NonNull
        @Getter(onMethod = @__(@ExcelExportObject(order = 0, headerName = "Name")))
        private String name;
        @NonNull
        @Getter(onMethod = @__(@ExcelExportObject(order = 1, headerName = "Number")))
        private Integer number;
        @NonNull
        @Getter(onMethod = @__(@ExcelExportObject(order = 2, headerName = "Decimal Number")))
        private Double decimal;
        @NonNull
        @Getter(onMethod = @__(@ExcelExportObject(order = 3, headerName = "Date with formatting")))
        private LocalDate date;
        @NonNull
        @Getter(onMethod = @__(@ExcelExportObject(order = 4, headerName = "Enum")))
        private SimpleEnum simpleEnum;
        @NonNull
        private Integer nonExportValue;
    }

    @ToString(callSuper = true)
    @Getter
    public static class SimpleClassWithList extends SimpleClass {
        @NonNull
        @Getter(onMethod = @__(@ExcelExportObject(order = 4, headerName = "Integers")))
        private List<Integer> integers;
        @NonNull
        @Getter(onMethod = @__(@ExcelExportObject(order = 4, headerName = "Strings")))
        private List<String> strings;

        public SimpleClassWithList(String name, Integer number, Double decimal, LocalDate date, SimpleEnum simpleEnum, Integer nonExportValue, List<Integer> integers, List<String> strings) {
            super(name, number, decimal, date, simpleEnum, nonExportValue);
            this.integers = integers;
            this.strings = strings;
        }
    }

    @AllArgsConstructor
    @ToString(callSuper = true)
    @Getter
    public static class SimpleClassListWrapper {
        @NonNull
        @Getter(onMethod = @__(@ExcelExportObject(order = 0, headerName = "Simple List")))
        private List<SimpleClass> simpleList;
    }
}