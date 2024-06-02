package com.maxexplode;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeFormatterBuilder;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

public class ExcelCustomDateFormat extends BaseFormat {

    /*
        Excel uses 1900 date system, that is why base year is set as 1899-12-30, any date specified in the
        excel row denotes to number of day after this date.
        Ex - If excel date is 44576, then the actual date is 1899-12-30 + 44576 days.
     */
    //TODO Check for leap year
    private final LocalDate baseExcelDate = LocalDate.of(1899, 12, 30);

    private static final Map<Integer, DateTimeFormatterBuilder> PREDEFINED_EXCEL_DATE_FMT = new HashMap<>();

    static {
        /*
        These are excel predefined formats
        ideal if any format id is less than 164 it is a standard format which is coming from Excel itself.
        At the moment added only 14 number format id which denotes to mm/dd/yyyy format.
        according to need please add any other date format to below map with the pattern.
        */
        PREDEFINED_EXCEL_DATE_FMT.put(14, new DateTimeFormatterBuilder()
                .appendPattern("[MM/dd[/yyyy]]"));
    }

    @Override
    public String format(Integer formatId, String value) {
        LocalDate localDate = baseExcelDate.plusDays(Long.parseLong(value));
        if (null == currentFormat) {
            DateTimeFormatterBuilder pattern = PREDEFINED_EXCEL_DATE_FMT.get(formatId);
            if (null == pattern) {
                throw new RuntimeException("Unknown date type : " + formatId);
            }
            return pattern.toFormatter().format(localDate);
        } else {
            return DateTimeFormatter.ofPattern(currentFormat).format(localDate);
        }
    }

    @Override
    public boolean supports(int formatId, String format) {
        if (PREDEFINED_EXCEL_DATE_FMT.containsKey(formatId) | DateUtil.isADateFormat(formatId, format)) {
            sanitizeAndSet(format);
            return true;
        }
        return false;
    }

    @Override
    public Set<Integer> supportedFormats() {
        return PREDEFINED_EXCEL_DATE_FMT.keySet();
    }

    public void sanitizeAndSet(String format) {
        String[] formatChunks = format.split(";");
        setCurrentFormat(formatChunks[0].replaceAll("m", "M"));
    }
}