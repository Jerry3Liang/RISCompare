package com.uitc;

import java.util.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

public class DateConverter {

    public static List<String> convertDates(String input) {
        String[] dates = input.split("\\|");
        Map<String, Integer> counter = new HashMap<>();
        List<String> result = new ArrayList<>();

        for (String date : dates) {

            //先去掉 "/"
            String cleanDate = date.replace("/", "");
            //計算同日期次數
            int count = counter.getOrDefault(cleanDate, 0) + 1;
            counter.put(cleanDate, count);
            //加上序號
            result.add(count + "_" + cleanDate);
        }

        //按照日期排序
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMdd");
        result.sort(Comparator.comparing(s -> LocalDate.parse(s.split("_")[1], formatter)));

        return result;
    }
}
