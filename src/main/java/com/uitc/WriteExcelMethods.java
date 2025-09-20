package com.uitc;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class WriteExcelMethods {

    public static Boolean exportIntegrationXlsx(
            String exportPath,
            Map<String, Integer> correctMap,
            Map<String, List<Object>> needUpdatedMap,
            Map<String, List<Object>> failureMap,
            List<String> noDownloadReportList,
            List<String> expYearReportIdList,
            String excelExportPartName
    ) {
        //產生當前時間 yyyy年MM月dd日
        String generateTime = DatetimeConverter.getSYSTime(4);

        //設定Excel表頭
        List<List<String>> correctHeader = new ArrayList<>();
        List<List<String>> needUpdatedHeader = new ArrayList<>();
        List<List<String>> failureHeader = new ArrayList<>();
        List<List<String>> noDownloadHeader = new ArrayList<>();
        List<List<String>> expYearHeader = new ArrayList<>();

        //欄位名稱
        correctHeader.add(Arrays.asList("報表代號", "總份數"));
        needUpdatedHeader.add(Arrays.asList("報表代號", "總份數", "原本就正確", "需校正", "有問題報表", "校正後正確", "需補檔", "最後需補檔日期", "有問題報表檔名", "需校正檔案部分名稱"));
        failureHeader.add(Arrays.asList("報表代號", "總份數", "與 OnDemand 差異", "備註"));
        noDownloadHeader.add(List.of("報表代號"));
        expYearHeader.add(List.of("報表代號"));

        try{
            createIntegrationXlsxFile(
                    exportPath,
                    correctHeader, needUpdatedHeader, failureHeader, noDownloadHeader, expYearHeader,
                    correctMap, needUpdatedMap, failureMap, noDownloadReportList, expYearReportIdList,
                    generateTime,
                    "Times New Roman",
                    excelExportPartName
            );

            return true;
        } catch (Exception e) {
            return false;
        }
    }

    private static void createIntegrationXlsxFile(
            String exportPath,
            List<List<String>> correctHeader,
            List<List<String>> needUpdatedHeader,
            List<List<String>> failureHeader,
            List<List<String>> noDownloadHeader,
            List<List<String>> expYearHeader,
            Map<String, Integer> correctMap,
            Map<String, List<Object>> needUpdatedMap,
            Map<String, List<Object>> failureMap,
            List<String> noDownloadReportList,
            List<String> expYearReportIdList,
            String generateTime,
            String fontName,
            String excelExportPartName
    ) throws IOException {
        SXSSFWorkbook workbook = new SXSSFWorkbook(100);
        SXSSFSheet correctSheet = workbook.createSheet("完全正確");
        SXSSFSheet updatedSheet = workbook.createSheet("需校正及補檔");
        SXSSFSheet failureSheet = workbook.createSheet("數量不對需重新下載");
        SXSSFSheet noDownloadSheet = workbook.createSheet("沒有下載到報表");
        SXSSFSheet expYearSheet = workbook.createSheet("過期報表");

        Map<Integer, Double> columnWidthMultiplier = new HashMap<>();
        columnWidthMultiplier.put(0, 1.3);
        columnWidthMultiplier.put(1, 1.3);
        columnWidthMultiplier.put(2, 1.3);
        columnWidthMultiplier.put(3, 1.3);// 第 3 列用 2.5 倍

        createHeaderAndStyleForExcel(correctHeader, workbook, correctSheet, fontName, columnWidthMultiplier);
        createHeaderAndStyleForExcel(needUpdatedHeader, workbook, updatedSheet, fontName, columnWidthMultiplier);
        createHeaderAndStyleForExcel(failureHeader, workbook, failureSheet, fontName, columnWidthMultiplier);
        createHeaderAndStyleForExcel(noDownloadHeader, workbook, noDownloadSheet, fontName, columnWidthMultiplier);
        createHeaderAndStyleForExcel(expYearHeader, workbook, expYearSheet, fontName, columnWidthMultiplier);

        CellStyle contentStyle = createContentAndStyleForExcel(workbook, fontName);

        //將 data 寫入 Excel (Sheet : 完全正確)
        writeMapToSheet(correctSheet, correctMap, contentStyle, true);

        //將 data 寫入 Excel (Sheet : 需校正及補檔)
        writeMapToSheet(updatedSheet, needUpdatedMap, contentStyle, true);

        //將 data 寫入 Excel (Sheet : 數量不對需重新下載)
        writeMapToSheet(failureSheet, failureMap, contentStyle, false);

        //將 data 寫入 Excel (Sheet : 沒有下載到報表)
        writeListToSheet(noDownloadSheet, noDownloadReportList, contentStyle);

        //將 data 寫入 Excel (Sheet : 過期報表)
        writeListToSheet(expYearSheet, expYearReportIdList, contentStyle);

        //寫出檔案
        FileOutputStream fileOut = new FileOutputStream(exportPath + "/" + excelExportPartName + "_Integration_Data_Result_" + generateTime + ".xlsx");
        workbook.write(fileOut);
        workbook.close();
        fileOut.close();
    }

    public static Boolean exportLackReportXlsx(
            String exportPath,
            Map<String, List<Object>> excel2ndExistMap,
            Map<String, List<Object>> excel2ndNotExistMap,
            String excelExportPartName
    ) {
        //產生當前時間 yyyy年MM月dd日
        String generateTime = DatetimeConverter.getSYSTime(4);

        //設定Excel表頭
        List<List<String>> existHeader = new ArrayList<>();
        List<List<String>> notExistHeader = new ArrayList<>();

        //欄位名稱
        existHeader.add(Arrays.asList("報表代號", "已存在報表份數", "原本就正確", "需校正", "有問題報表", "校正後正確", "需補檔", "最後需補檔日期", "有問題報表檔名", "需校正檔案部分名稱"));
        notExistHeader.add(Arrays.asList("報表代號", "所缺報表數量", "最後需補檔日期", "所缺檔案部分名稱"));

        try{
            createLackReportXlsxFile(
                    exportPath,
                    existHeader, notExistHeader,
                    excel2ndExistMap, excel2ndNotExistMap,
                    generateTime,
                    "Times New Roman",
                    excelExportPartName
            );

            return true;
        } catch (Exception e) {
            return false;
        }
    }

    private static void createLackReportXlsxFile(
            String exportPath,
            List<List<String>> existHeader,
            List<List<String>> notExistHeader,
            Map<String, List<Object>> excel2ndExistMap,
            Map<String, List<Object>> excel2ndNotExistMap,
            String generateTime,
            String fontName,
            String excelExportPartName
    ) throws IOException {
        SXSSFWorkbook workbook = new SXSSFWorkbook(100);
        SXSSFSheet existSheet = workbook.createSheet("已存在");
        SXSSFSheet notExistSheet = workbook.createSheet("不存在");

        Map<Integer, Double> columnWidthMultiplier = new HashMap<>();
        columnWidthMultiplier.put(0, 1.3);
        columnWidthMultiplier.put(1, 1.3);
        columnWidthMultiplier.put(2, 1.3);
        columnWidthMultiplier.put(3, 1.3);// 第 3 列用 2.5 倍

        createHeaderAndStyleForExcel(existHeader, workbook, existSheet, fontName, columnWidthMultiplier);
        createHeaderAndStyleForExcel(notExistHeader, workbook, notExistSheet, fontName, columnWidthMultiplier);

        CellStyle contentStyle = createContentAndStyleForExcel(workbook, fontName);

        //將 data 寫入 Excel (Sheet : 已存在)
        writeMapToSheet(existSheet, excel2ndExistMap, contentStyle, false);

        //將 data 寫入 Excel (Sheet : 不存在")
        writeMapToSheet(notExistSheet, excel2ndNotExistMap, contentStyle, false);

        //寫出檔案
        FileOutputStream fileOut = new FileOutputStream(exportPath + "/" + excelExportPartName + "_Lack_Report_Data_Result_" + generateTime + ".xlsx");
        workbook.write(fileOut);
        workbook.close();
        fileOut.close();
    }

    /**
     * 所有 Excel 表頭部分的通用樣式
     * @param header: 表頭內容陣列
     * @param workbook: .xlsx 的 Excel
     * @param sheet: Excel 的內頁
     * @param columnWidthMultiplier: 動態設定列寬的索引與倍數
     */
    private static void createHeaderAndStyleForExcel(
            List<List<String>> header,
            SXSSFWorkbook workbook,
            SXSSFSheet sheet,
            String fontName,
            Map<Integer, Double> columnWidthMultiplier //動態設定列寬的索引與倍數
    ){
        //設定自適應列寬
        sheet.trackAllColumnsForAutoSizing();

        byte[] rgb = new byte[] {
                (byte) 255,
                (byte) 255,
                (byte) 255 };

        Font font = workbook.createFont();
        font.setFontName(fontName);

        XSSFFont headerFont = (XSSFFont) workbook.createFont();
        headerFont.setColor(new XSSFColor(rgb, null));
        headerFont.setFontName(fontName);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setBold(true);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setBorderBottom(BorderStyle.MEDIUM);
        headerStyle.setBorderTop(BorderStyle.MEDIUM);
        headerStyle.setBorderRight(BorderStyle.MEDIUM);
        headerStyle.setBorderLeft(BorderStyle.MEDIUM);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setFont(headerFont);

        //背景顏色
        headerStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        //設定表頭
        for (int i = 0; i < header.size(); i++) {
            Row row = sheet.createRow(i);

            for (int j = 0; j < header.get(i).size(); j++) {
                Cell cell = row.createCell(j);
                cell.setCellStyle(headerStyle);
                cell.setCellValue(header.get(i).get(j)==null ? "" : header.get(i).get(j));
                //設定自適應列寬
                sheet.autoSizeColumn(j);

                //檢查是否有特定列需要使用倍數調整寬度
                if (columnWidthMultiplier.containsKey(j)) {
                    double multiplier = columnWidthMultiplier.get(j);
                    sheet.setColumnWidth(j, (int) (sheet.getColumnWidth(j) * multiplier));
                } else {
                    sheet.setColumnWidth(j, (sheet.getColumnWidth(j) * 12 / 10));
                }
            }
        }
    }

    /**
     * 所有 Excel 內容部分的通用樣式
     * @param workbook: .xlsx 的 Excel
     * @return : CellStyle
     */
    private static CellStyle createContentAndStyleForExcel(SXSSFWorkbook workbook, String fontName){

        CellStyle contentStyle = workbook.createCellStyle();
        contentStyle.setAlignment(HorizontalAlignment.CENTER);
        contentStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        Font contentFont = workbook.createFont();
        contentFont.setFontName(fontName);
        contentFont.setFontHeightInPoints((short) 12);
        contentStyle.setFont(contentFont);

        return contentStyle;
    }

    private static void writeMapToSheet(SXSSFSheet sheet, Map<String, ?> dataMap, CellStyle contentStyle, boolean needSorted) {
        List<String> keys = new ArrayList<>(dataMap.keySet());
        if(needSorted) {
            Collections.sort(keys);
        }

        for(int i = 0; i < keys.size(); i++) {
            String key = keys.get(i);
            Row row = sheet.createRow(i + 1);
            createCell(row, 0, key, contentStyle);

            Object value = dataMap.get(key);

            if(value instanceof List<?>) {
                List<?> list = (List<?>) value;
                for(int j = 0; j < list.size(); j++) {
                    createCell(row, j + 1, list.get(j), contentStyle);
                }
            } else {
                createCell(row, 1, value, contentStyle);
            }

        }
    }

    private static void writeListToSheet(SXSSFSheet sheet, List<String> dataList, CellStyle contentStyle) {
        for(int i = 0; i < dataList.size(); i++) {
            Row row = sheet.createRow(i + 1);
            createCell(row, 0, dataList.get(i), contentStyle);
        }
    }

    /**
     * 設置 Excel 儲存格的通用方法
     * @param row: 已創建好的 row
     * @param columnIndex: 欲寫入 Excel 的 Column 索引
     * @param value: 欲寫入 Excel 的值
     * @param style: 欲使用的樣式
     */
    private static void createCell(Row row, int columnIndex, Object value, CellStyle style) {

        Cell cell = row.createCell(columnIndex);
        if(value instanceof String) {
            cell.setCellValue((String) value);
        } else if(value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if(value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if(value instanceof List) {
            @SuppressWarnings("unchecked")
            List<Object> list = (List<Object>) value;
            String joined = list.stream()
                    .map(String::valueOf)//確保轉成字串
                    .reduce((a, b) -> a + ", " + b)  // 用逗號串接
                    .orElse("");
            cell.setCellValue(joined);
        } else if (value instanceof Set) {
            //假設 Set 裡的元素型別不固定，全部轉字串
            @SuppressWarnings("unchecked")
            Set<Object> set = (Set<Object>) value;
            String joined = set.stream()
                    .map(String::valueOf)
                    .reduce((a, b) -> a + ", " + b)
                    .orElse("");
            cell.setCellValue(joined);
        }

        cell.setCellStyle(style);
    }
}
