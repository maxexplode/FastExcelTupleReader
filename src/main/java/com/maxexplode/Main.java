package com.maxexplode;

import lombok.Data;

import javax.xml.stream.XMLStreamException;
import java.io.IOException;
import java.net.URL;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class Main {
    public static void main(String[] args) throws XMLStreamException, IOException {
        URL resource = Main.class.getResource("/sample_data_small.xlsx");
        assert resource != null;
        FastExcelTupleReader<Person> fastExcelTupleReader = new FastExcelTupleReader<>(
                resource.getPath(), Person.class, ReadOptions.builder()
                .sheetIdx(1)
                .dataRowIdx("1")
                .build()
        );

        Stream<Person> personStream = fastExcelTupleReader.read();

        List<Person> personList = personStream.collect(Collectors.toList());

        System.out.println("Sucessfully read : " + personList.size());
    }

    @Data
    @ExcelRow
    static class Person{
        @ExcelCellName
        private String name;
        @ExcelCellName
        private String description;
        @ExcelCellName
        private String address;
        @ExcelCellName
        private String phoneno;
        @ExcelCellName
        private Date dob;
    }
}

