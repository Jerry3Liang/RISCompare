package com.uitc;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.util.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class DateConverter {

    private static final Logger logger = LoggerFactory.getLogger(DateConverter.class);

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
            String reportName,
            boolean isEArea
    ) throws IOException {

        Map<String, String> specialPatternMap = config.getPatternByReport();
        // 建立正則：關鍵字 + 任意 0~20 個字元 + 日期 (粗略)
        String keywordPattern = String.join("|", config.getKeywords());
        Pattern regex = Pattern.compile(
                "(?:" +
                        "(" + keywordPattern + ")[:：]?\\s{0,4}" +
                       ")?" +
                        "(" +
                            "\\d{4}[./-]\\d{1,2}[./-]\\d{1,2}" +
                            "|" +
                            "\\d{1,2}[./-]\\d{1,2}[./-]\\d{2,4}" +
                            "|" +
                            "\\d{4}\\s*年\\s*\\d{1,2}\\s*月\\s*\\d{1,2}\\s*日" +
                        ")"
        );

        String keyWord;
        String dateStr;
        Matcher matcher = regex.matcher(line);
        //Matcher 的 find() 方法是每呼叫一次，就往下一個匹配移動，所以第一次 mach 到在呼叫一次就會 false
        if(matcher.find()) {
            keyWord = matcher.group(1); //group 1 是 keyWord 部分
            dateStr = matcher.group(2); //group 2 是日期部分

            if (keyWord == null) {
                String logMessage = isEArea ? "E-2-3. 沒有 keyword，只抓到日期！" : "G-2-3. 沒有 keyword，只抓到日期！";
                logger.info(logMessage);
            } else {
                String logMessage1 = isEArea ? "E-2-3. 內文比對辨識到的關鍵字 ： {}" : "G-2-3. 內文比對到的關鍵字 ： {}";
                logger.info(logMessage1, keyWord);
            }

            String logMessage2 = isEArea ? "E-2-4. 內文比對辨識到的日期 ： {}" : "G-2-4. 內文比對到的日期 ： {}";
            logger.info(logMessage2, dateStr);

            // ✅ 防呆：一定要有分隔符
            if(!dateStr.matches(".*[./\\-年月].*")) {
                return "無法辨識格式！";
            }

            LocalDate parsedDate = parseDateFlexible(dateStr, specialPatternMap, reportName);
            if (parsedDate != null) {
                return parsedDate.format(DateTimeFormatter.ofPattern("yyyyMMdd"));
            }
        }

        return "無法辨識格式！！！";
    }

    private static LocalDate  parseDateFlexible(
            String dateStr,
            Map<String, String> specialPatternMap,
            String reportName
    ) {
        //先把中文年月日替換成 /
        String cleanDate = dateStr.replace("年","/").replace("月","/").replace("日","");

        //清掉空白
        cleanDate = cleanDate.replaceAll("\\s+", "");

        //1. 先判斷完整名稱
        if(specialPatternMap.containsKey(reportName)) {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern(specialPatternMap.get(reportName));
            return LocalDate.parse(cleanDate, formatter);
        } else { //2. 再判斷開頭
            for(String specialPattern : specialPatternMap.keySet()) {
                if(reportName.startsWith(specialPattern)) {
                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern(specialPatternMap.get(specialPattern));
                    return LocalDate.parse(cleanDate, formatter);
                }
            }
        }

        String[] parts = cleanDate.split("[/.-]");

        if (parts.length != 3) {
            return null;
        }

        int year, month, day;
        // ✅ yyyy/MM/dd 或 民國
        if (parts[0].length() == 4) {
            year = Integer.parseInt(parts[0]);
            month = Integer.parseInt(parts[1]);
            day = Integer.parseInt(parts[2]);

            // 👉 只在「確定是年」時才轉民國
            if (year < 1911) {
                year += 1911;
            }
        } else if (parts[2].length() == 2) {
            // ✅ MM/dd/yy
            month = Integer.parseInt(parts[0]);
            day = Integer.parseInt(parts[1]);
            year = 2000 + Integer.parseInt(parts[2]);
        } else if (parts[2].length() == 4) {
            // ✅ MM/dd/yyyy
            month = Integer.parseInt(parts[0]);
            day = Integer.parseInt(parts[1]);
            year = Integer.parseInt(parts[2]);
        } else {
            return null;
        }

        return LocalDate.of(year, month, day);
//        for (String pattern : datePatterns) {
//
//            try {
//                DateTimeFormatter formatter = DateTimeFormatter.ofPattern(pattern);
//                return LocalDate.parse(cleanDate, formatter);
//            } catch (DateTimeParseException ignored) {}
//        }
//
//        return null;
    }
}
