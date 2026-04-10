package com.uitc;

import java.io.IOException;
import java.time.format.DateTimeParseException;
import java.util.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class DateConverter {

    public static List<String> convertDates(String input) {
        System.out.println("input : " + input);
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

        System.out.println("result : " + result);
        //按照日期排序
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMdd");
        result.sort(Comparator.comparing(s -> LocalDate.parse(s.split("_")[1], formatter)));

        return result;
    }

    public static String extractDatesFromFile(
            String line,
            ConfigLoader config,
            String reportName
    ) throws IOException {

        Map<String, String> specialPatternMap = config.getPatternByReport();
        // 建立正則：關鍵字 + 任意 0~20 個字元 + 日期 (粗略)
        String keywordPattern = String.join("|", config.getKeywords());
        Pattern regex = Pattern.compile("(" + keywordPattern + "){0,2}(\\d{1,4}[./-年]\\d{1,2}[./-月]\\d{1,4}日?)");
        String dateStr = null;
        Matcher matcher = regex.matcher(line);
        if(matcher.find()) {
            dateStr = matcher.group(2); // group 2 是日期部分
            LocalDate parsedDate = parseDateFlexible(dateStr, config.getDatePatterns(), specialPatternMap, reportName);
            if (parsedDate != null) {
                return parsedDate.format(DateTimeFormatter.ofPattern("yyyyMMdd"));
            }
        }

        return "無法辨識格式: " + dateStr;
    }

    private static LocalDate  parseDateFlexible(
            String dateStr,
            List<String> datePatterns,
            Map<String, String> specialPatternMap,
            String reportName
    ) {
        // 先把中文年月日替換成 /
        String cleanDate = dateStr.replace("年","/").replace("月","/").replace("日","");

        // 1. 先判斷完整名稱
        if(specialPatternMap.containsKey(reportName)) {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern(specialPatternMap.get(reportName));
            return LocalDate.parse(cleanDate, formatter);
        }

        // 2. 再判斷開頭
        else {
            for(String specialPattern : specialPatternMap.keySet()) {
                if(reportName.startsWith(specialPattern)) {
                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern(specialPatternMap.get(specialPattern));
                    return LocalDate.parse(cleanDate, formatter);
                }
            }
        }

        for (String pattern : datePatterns) {
            try {
                DateTimeFormatter formatter = DateTimeFormatter.ofPattern(pattern);
                return LocalDate.parse(cleanDate, formatter);
            } catch (DateTimeParseException ignored) {}
        }

        return null;
    }
}
