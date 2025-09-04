package com.uitc;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class TransferMethods {

    public static ParamsContent getParamsContent(String paramsFilePath) {

        ParamsContent paramsContent = new ParamsContent();
        File file = new File(paramsFilePath);
        if(file.isFile() && isValidExcelFile(file)) {
            String filePathString = file.getAbsolutePath();
            System.out.println("讀取到的參數檔：" + file.getName());
            try (FileInputStream fileInputStream = new FileInputStream(filePathString)) {
                Workbook workbook = WorkbookFactory.create(fileInputStream);
                Sheet sheet = workbook.getSheetAt(0);

                Integer db2024ExcelFilePathPositionRowIndex = findRowIndexByCellValue(sheet, "DB 2024 Excel Path", 0);
                paramsContent.setDataBase2024ExcelPath(findValueByRowIndex(sheet, db2024ExcelFilePathPositionRowIndex));

                Integer folderPathPositionRowIndex = findRowIndexByCellValue(sheet, "Search Folder Path", 0);
                paramsContent.setSearchFolderPathURI(findValueByRowIndex(sheet, folderPathPositionRowIndex));

                Integer reportInYearRowIndex = findRowIndexByCellValue(sheet, "Report In Year", 0);
                paramsContent.setReportInYear(findValueByRowIndex(sheet, reportInYearRowIndex));

                Integer updatedFileFolderPathPositionRowIndex = findRowIndexByCellValue(sheet, "Updated File Folder Path", 0);
                paramsContent.setUpdatedFileFolderPathURI(findValueByRowIndex(sheet, updatedFileFolderPathPositionRowIndex));

                Integer compareResultExportPathRowIndex = findRowIndexByCellValue(sheet, "Compare Result Export Path", 0);
                paramsContent.setCompareResultExportPath(findValueByRowIndex(sheet, compareResultExportPathRowIndex));

                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        return paramsContent;
    }

    public static Map<String, Double> get2024ReportCounts(String db2024ExcelPath) {

        Map<String, Double> reportNameVSCount = new HashMap<>();
        File file = new File(db2024ExcelPath);
        if(file.isFile() && isValidExcelFile(file)) {
            String filePathString = file.getAbsolutePath();
            try (FileInputStream fileInputStream = new FileInputStream(filePathString)) {
                Workbook workbook = WorkbookFactory.create(fileInputStream);
                Sheet sheet = workbook.getSheetAt(0);
                double total = 0.0;
                for(int rowIndex = 1; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    Cell reportIdCell = row.getCell(0);
                    Cell totalCountCell = row.getCell(1);
                    String reportId =  getStringCellValue(reportIdCell);
                    if (totalCountCell != null && totalCountCell.getCellType() == CellType.NUMERIC) {
                        total = totalCountCell.getNumericCellValue();
                    }
                    reportNameVSCount.put(reportId, total);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        return reportNameVSCount;
    }

    //參數檔用
    private static Integer findRowIndexByCellValue(Sheet sheet, String keyWord, Integer columnIndex) {

        for (Row row : sheet) {
            if (row == null) continue;
            Cell cell = row.getCell(columnIndex);
            if (cell != null && cell.getCellType() == CellType.STRING) {
                String cellValue = cell.getStringCellValue();
                if (cellValue.contains(keyWord)) {
                    return row.getRowNum(); // 返回匹配關鍵字的 row index
                }
            }
        }

        return null;
    }

    private static String findValueByRowIndex(Sheet sheet, Integer rowIndex) {

        Row row = sheet.getRow(rowIndex);
        Cell valueCell = row.getCell(1);

        if(valueCell.getCellType() == CellType.STRING) {
            return getStringCellValue(valueCell);
        } else if(valueCell.getCellType() == CellType.NUMERIC) {
            return getNumericCellValue(valueCell);
        }

        return null;
    }

    public static Boolean exportXlsx(
            String exportPath,
            Map<String, Integer> correctMap,
            Map<String, List<Object>> needUpdatedMap,
            Map<String, List<Object>> failureMap,
            String excelExportPartName
    ) {
        //產生當前時間 yyyy年MM月dd日
        String generateTime = DatetimeConverter.getSYSTime(4);

        //設定Excel表頭
        List<List<String>> correctHeader = new ArrayList<>();
        List<List<String>> needUpdatedHeader = new ArrayList<>();
        List<List<String>> failureHeader = new ArrayList<>();

        //欄位名稱
        correctHeader.add(Arrays.asList("報表代號", "總份數"));
        needUpdatedHeader.add(Arrays.asList("報表代號", "總份數", "原本就正確", "需校正", "有問題報表", "校正後正確", "需補檔", "最後需補檔日期", "有問題報表檔名", "需校正檔案部分名稱"));
        failureHeader.add(Arrays.asList("報表代號", "總份數", "備註"));

        try{
            createCWaveXlsxFile(
                    exportPath,
                    correctHeader, needUpdatedHeader, failureHeader,
                    correctMap, needUpdatedMap, failureMap,
                    generateTime,
                    "Times New Roman",
                    excelExportPartName
            );

            return true;
        } catch (Exception e) {
            return false;
        }
    }

    private static void createCWaveXlsxFile(
            String exportPath,
            List<List<String>> correctHeader,
            List<List<String>> needUpdatedHeader,
            List<List<String>> failureHeader,
            Map<String, Integer> correctMap,
            Map<String, List<Object>> needUpdatedMap,
            Map<String, List<Object>> failureMap,
            String generateTime,
            String fontName,
            String excelExportPartName
    ) throws IOException {
        SXSSFWorkbook workbook = new SXSSFWorkbook(100);
        SXSSFSheet correctSheet = workbook.createSheet("完全正確");
        SXSSFSheet updatedSheet = workbook.createSheet("需校正及補檔");
        SXSSFSheet failureSheet = workbook.createSheet("需全部重新下載");

        Map<Integer, Double> columnWidthMultiplier = new HashMap<>();
        columnWidthMultiplier.put(0, 1.3);
        columnWidthMultiplier.put(1, 1.3);
        columnWidthMultiplier.put(2, 1.3);
        columnWidthMultiplier.put(3, 1.3);// 第 3 列用 2.5 倍

        createHeaderAndStyleForExcel(correctHeader, workbook, correctSheet, fontName, columnWidthMultiplier);
        createHeaderAndStyleForExcel(needUpdatedHeader, workbook, updatedSheet, fontName, columnWidthMultiplier);
        createHeaderAndStyleForExcel(failureHeader, workbook, failureSheet, fontName, columnWidthMultiplier);

        CellStyle contentStyle = createContentAndStyleForExcel(workbook, fontName);

        //設定內容
        List<String> correctMapKeys = new ArrayList<>(correctMap.keySet());
        Collections.sort(correctMapKeys);

        //將 data 寫入 Excel (Sheet : 完全正確)
        for (int i = 0; i < correctMapKeys.size(); i++) {

            String key = correctMapKeys.get(i);

            int value = correctMap.get(key);

            // 建立第二列 (Row)
            Row row = correctSheet.createRow(i + 1);
            createCell(row, 0, key, contentStyle);
            createCell(row, 1, value, contentStyle);
        }

        //設定內容
        List<String> needUpdatedMapKeys = new ArrayList<>(needUpdatedMap.keySet());
        Collections.sort(needUpdatedMapKeys);

        //將 data 寫入 Excel
        for (int i = 0; i < needUpdatedMapKeys.size(); i++) {

            String key = needUpdatedMapKeys.get(i);

            List<Object> values = needUpdatedMap.get(key);

            // 建立第二列 (Row)
            Row row = updatedSheet.createRow(i + 1);
            createCell(row, 0, key, contentStyle);

            for(int j = 0; j < values.size(); j++){
                createCell(row, (j + 1), values.get(j), contentStyle);
            }
        }

        //設定內容
        List<String> failureMapKeys = new ArrayList<>(failureMap.keySet());
        Collections.sort(failureMapKeys);

        //將 data 寫入 Excel (Sheet : 需全部重新下載)
        for (int i = 0; i < failureMapKeys.size(); i++) {

            String key = failureMapKeys.get(i);

            List<Object> values = failureMap.get(key);

            // 建立第二列 (Row)
            Row row = failureSheet.createRow(i + 1);
            createCell(row, 0, key, contentStyle);

            for(int j = 0; j < values.size(); j++){
                createCell(row, (j + 1), values.get(j), contentStyle);
            }
        }

        //寫出檔案
        FileOutputStream fileOut = new FileOutputStream(exportPath + "/" + excelExportPartName + "_Result_" + generateTime + ".xlsx");
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

    /**
     * 獲得讀取到的儲存格值。
     * @param cell: 所讀取的 Excel 檔裡的儲存格。
     * @return : 讀取的儲存格如為 String 資料型態就傳回該值，若不為 String 資料型態則傳回 null。
     */
    private static String getStringCellValue(Cell cell) {

        if (cell != null) {
            if (cell.getCellType() == CellType.STRING) {

                return cell.getStringCellValue().trim();
            }
        }

        return null; //Return Null for non-string cells or errors
    }

    /**
     * 獲得讀取到的儲存格值。
     * @param cell: 所讀取的 Excel 檔裡的儲存格。
     * @return : 讀取的儲存格如為 double 資料型態就傳回該值，若不為 double 資料型態則傳回 0.0。
     */
    private static String getNumericCellValue(Cell cell) {
        if (cell != null && cell.getCellType() == CellType.NUMERIC) {
            double numericValue = cell.getNumericCellValue();
            //檢查是否為整數
            if (numericValue == Math.floor(numericValue)) {
                return String.valueOf((long) numericValue);
            } else {
                return String.valueOf(numericValue);
            }
        }

        return null;
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

    /**
     * 檢查檔案名稱是否以 ".xlsx" 或 ".xls" 結尾 (符合 Excel 檔的副檔名)，及過濾掉檔名開頭是 ._ 或 ~$。
     * @param file：所有讀取到的檔案。
     * @return boolean: true (是 Excel 檔) 或 false (不是 Excel 檔)。
     */
    public static boolean isValidExcelFile(File file) {
        //檢查檔案名稱的副檔名是否為 ".xlsx" 或 ".xls"。
        if (!(file.getName().endsWith(".xlsx") || file.getName().endsWith(".xls"))) {
            return false;
        }

        // 檢查檔案名稱是否以 "._" 或 "~$" 開頭。
        if (file.getName().startsWith("._") || file.getName().startsWith("~$")) {
            return false;
        }
        return true;
    }

    /**
     * 檢查檔案名稱是否以 ".txt" 或 ".TXT" 結尾 (符合 Excel 檔的副檔名)，及過濾掉檔名開頭是 ._ 或 ~$ 或 .DS。
     * @param file：所有讀取到的檔案。
     * @return boolean: true (是 txt 檔) 或 false (不是 txt 檔)。
     */
    public static boolean isValidFile(File file) {
        //檢查檔案名稱的副檔名是否為 ".txt" 或 ".TXT"。
//        if (!(file.getName().endsWith(".txt") || file.getName().endsWith(".TXT"))) {
//            return false;
//        }

        // 檢查檔案名稱是否以 "._" 或 "~$" 或 ".DS"開頭。
        if (file.getName().startsWith("._") || file.getName().startsWith("~$") || file.getName().startsWith(".DS")) {
            return false;
        }
        return true;
    }
}
