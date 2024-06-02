package com.maxexplode;

import java.util.regex.Pattern;

public class DateUtil {
    
    // Function to check if a given format ID and format string represent a date format
    public static boolean isADateFormat(int formatId, String format) {
        // Standard date format IDs in Excel
        if (isStandardDateFormatId(formatId)) {
            return true;
        }

        // Check for custom date format strings
        if (format != null && !format.isEmpty()) {
            return isDateFormatString(format);
        }

        return false;
    }

    private static boolean isStandardDateFormatId(int formatId) {
        // Common date format IDs in Excel
        int[] dateFormatIds = {14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47};
        for (int id : dateFormatIds) {
            if (formatId == id) {
                return true;
            }
        }
        return false;
    }

    private static boolean isDateFormatString(String format) {
        // Regex pattern to match typical date formats in Excel
        String datePattern = "(?i).*[dy].*"; // Case insensitive match for 'd' or 'y'

        // Remove unwanted characters that might appear in date formats
        format = format.replaceAll("[\\s\"\\\\,\\*]", "");

        // Match against the date pattern
        return Pattern.matches(datePattern, format);
    }
}