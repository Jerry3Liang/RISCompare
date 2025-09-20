package com.uitc;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ReadExcelMethods {

    public static ParamsContent getParamsContent(String paramsFilePath) {

        ParamsContent paramsContent = new ParamsContent();
        File file = new File(paramsFilePath);
        if(file.isFile() && isValidExcelFile(file)) {
            String filePathString = file.getAbsolutePath();
            System.out.println("讀取到的參數檔：" + file.getName());
            try (FileInputStream fileInputStream = new FileInputStream(filePathString)) {
                Workbook workbook = WorkbookFactory.create(fileInputStream);
                Sheet sheet = workbook.getSheetAt(0);

                Integer db2024ExcelFilePathPositionRowIndex = findRowIndexByCellValue(sheet, "Year Excel Path", 0);
                paramsContent.setYearExcelPath(findValueByRowIndex(sheet, db2024ExcelFilePathPositionRowIndex));

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

    public static Map<String, List<Double>> getReportCountsAndExpYearForYear(String yearExcelPath) {

        Map<String, List<Double>> reportNameVSCountAndExpYear = new HashMap<>();
        File file = new File(yearExcelPath);
        if(file.isFile() && isValidExcelFile(file)) {
            String filePathString = file.getAbsolutePath();
            try (FileInputStream fileInputStream = new FileInputStream(filePathString)) {
                Workbook workbook = WorkbookFactory.create(fileInputStream);
                Sheet sheet = workbook.getSheetAt(0);
                double total;
                double expYear;
                for(int rowIndex = 1; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
                    List<Double> countAndExpYear = new ArrayList<>();
                    Row row = sheet.getRow(rowIndex);
                    Cell reportIdCell = row.getCell(0);
                    Cell totalCountCell = row.getCell(1);
                    Cell expYearCell = row.getCell(2);
                    String reportId =  getStringCellValue(reportIdCell);
                    if (totalCountCell != null && totalCountCell.getCellType() == CellType.NUMERIC) {
                        total = totalCountCell.getNumericCellValue();
                        countAndExpYear.add(total);
                    }

                    if (expYearCell != null && expYearCell.getCellType() == CellType.NUMERIC) {
                        expYear = expYearCell.getNumericCellValue();
                        countAndExpYear.add(expYear);
                    }

                    reportNameVSCountAndExpYear.put(reportId, countAndExpYear);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        return reportNameVSCountAndExpYear;
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
        return !file.getName().startsWith("._") && !file.getName().startsWith("~$");
    }
}
