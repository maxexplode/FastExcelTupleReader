package com.maxexplode;

import com.maxexplode.stereotype.ExcelCellName;
import com.maxexplode.stereotype.ExcelRow;
import lombok.Data;

import javax.xml.stream.XMLStreamException;
import java.io.IOException;
import java.net.URL;
import java.util.Date;
import java.util.List;
import java.util.stream.Stream;


public class Main {
    public static void main(String[] args) throws XMLStreamException, IOException {
        try {
            URL resource = Main.class.getResource("/sample_data_small.xlsx");
            assert resource != null;
            FastExcelTupleReader<Person> fastExcelTupleReader = new FastExcelTupleReader<>(
                    resource.getPath(), Person.class, ReadOptions.builder()
                    .sheetIdx(1)
                    .dataRowIdx("1")
                    .build()
            );

            Stream<Person> personStream = fastExcelTupleReader.read();

            List<Person> personList = personStream.toList();

            System.out.println("Successfully read : " + personList.size());
        }catch (Exception e){
            e.printStackTrace();
        }
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

