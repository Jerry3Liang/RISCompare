package com.uitc;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.*;
import java.util.stream.Collectors;

import static com.uitc.TransferMethods.get2024ReportCounts;

public class Main {

    private static final Logger logger = LoggerFactory.getLogger(Main.class);

    public static void main(String[] args) {
        //Mac 用
//        final String PARAMS_PATH = "/Users/liangchejui/Desktop/程式語言相關/RIS 比對小程式資料/參數檔/params_for_RIS.xlsx";
        //Windows 用
        final String PARAMS_PATH = "C:/比對小程式/Config/params_for_RIS.xlsx";
        logger.info("A. 開始讀取參數檔，參數檔 Path： {}", PARAMS_PATH);

        //讀取參數檔
        ParamsContent paramsContent = TransferMethods.getParamsContent(PARAMS_PATH);

        //取得 2024年報表代號對應總份數的 Map
        String db2024ExcelPath = paramsContent.getDataBase2024ExcelPath();
        Map<String, Double> db2024ExcelMap = get2024ReportCounts(db2024ExcelPath);
        System.out.println("2024年報表代號對應總份數 Excel 所在檔案路徑：" + db2024ExcelPath);
        logger.info("B. 2024年報表代號對應總份數 Excel 所在檔案路徑： {}", db2024ExcelPath);

//        System.out.println("db2024ExcelMap : " + db2024ExcelMap);

        //要比對的年份
        String searchYear = paramsContent.getReportInYear();
        logger.info("C. 比對報表年份： {}", searchYear);
        String searchFolderPathUrl = paramsContent.getSearchFolderPathURI();
        logger.info("D. 報表所在資料夾路徑： {}", searchFolderPathUrl);
        File currentFolders = new File(searchFolderPathUrl);
        File[] reportFolders = currentFolders.listFiles(File::isDirectory);
        assert reportFolders != null;

        //最後匯出 Excel 需要的資料
        //1. 數量正確且不需校正
        Map<String, Integer> correctMap = new HashMap<>();

        //2. 數量正確但需要校正
        Map<String, List<Object>> needUpdatedMap = new HashMap<>();

        //3. 數量不正確，可能需全部重新下載
        Map<String, List<Object>> failureMap = new HashMap<>();

        //遍歷所有報表代號的資料夾
        for (File reportFolder : reportFolders) {
            //計算讀取報表內容行數
            int countLine = 0;

            //判斷是否為英文報表
            boolean isEnglishReport = false;

            //2. 的欄位資料
            List<Object> needUpdatedColumnList = new ArrayList<>();
            int totalCount;
            int originalCorrectCount = 0;
            int needUpdate = 0;
            int updatedCorrectCount;
            int failureCount = 0;
            //建立全部檔案名稱日期 Set
            Set<String> dateSet = new HashSet<>();
            //建立錯誤報表檔案名稱 List
            List<String> failureFileName = new ArrayList<>();
            //建立日期原本就正確 Set
            Set<String> originalCorrectDate = new HashSet<>();
            //建立需校正名單
            Set<String> needUpdateDate = new HashSet<>();
            //建立校正完名稱 Set，儲存改完名的日期
            Set<String> UpdatedDate = new HashSet<>();

            //第 3. 的欄位資料
            List<Object> failuredColumnList = new ArrayList<>();

            //取得報表資料夾的名稱
            String reportFolderName = reportFolder.getName();
            System.out.println("報表資料夾 : " + reportFolderName);
            logger.info("D-1. 報表資料夾： {}", reportFolderName);

            //報表代號的資料夾其報表代號有在 2024年報表代號對應總份數的 Map
            if(!db2024ExcelMap.containsKey(reportFolderName)) {
                System.out.println(reportFolderName + " 不存在");
                logger.info("D-1. 報表代號： {} 不存在於 2024年 Excel！", reportFolder);
                continue;
            }

            File[] reportFolderFiles = reportFolder.listFiles(file -> file.isFile() && !file.isHidden());
            assert reportFolderFiles != null;

            //建立檔案名稱的 List
            List<String> fileNameList = new ArrayList<>();

            for(File reportFile : reportFolderFiles) {
                String reportFileName = reportFile.getName();
                String reportFileNameYear = reportFileName.substring(reportFileName.lastIndexOf("_") + 1, reportFileName.lastIndexOf("_") + 5);
                if(Objects.equals(searchYear, reportFileNameYear)) {
                    fileNameList.add(reportFileName);
                }
            }

            System.out.println("fileNameList : " + fileNameList);

            double currentCompare2024Count = Math.abs(db2024ExcelMap.get(reportFolderName) - fileNameList.size());

            if(fileNameList.size() == db2024ExcelMap.get(reportFolderName)) {
                totalCount = fileNameList.size();
                logger.info("E. 報表資料夾： {} 與 RIS 系統上數量一樣！", reportFolder);

                //遍歷報表代號的資料夾裡所有檔案並將檔案名稱的日期存取到 dateSet
                for(File reportFile : reportFolderFiles) {
                    String reportFileName = reportFile.getName();
                    String reportFileNameYear = reportFileName.substring(reportFileName.lastIndexOf("_") + 1, reportFileName.lastIndexOf("_") + 5);
                    if(Objects.equals(searchYear, reportFileNameYear)) {
                        String fileNameDate = reportFileName.substring(reportFileName.indexOf("_") + 1, reportFileName.lastIndexOf("_") + 9);
                        dateSet.add(fileNameDate);
                    }
                }

                //遍歷報表代號的資料夾裡所有檔案
                for (File reportFile : reportFolderFiles) {
                    if(TransferMethods.isValidFile(reportFile)) {
                        String reportFileName = reportFile.getName();
                        String reportFileNameYear = reportFileName.substring(reportFileName.lastIndexOf("_") + 1, reportFileName.lastIndexOf("_") + 5);
                        if(Objects.equals(searchYear, reportFileNameYear)) {
                            String reportFileNameDate = reportFileName.substring(reportFileName.indexOf("_") + 1, reportFileName.lastIndexOf("_") + 9);

                            //Windows 電腦，檔案如果為空白，容量為 0
                            if(reportFile.length() == 0) {
                                failureCount++;
                                failureFileName.add(reportFileName);
                                logger.debug("E-1-1. 檔案名稱 : {} 為空白報表", reportFileName);
                                continue;
                            }

                            logger.info("E-1-1. 檔案名稱： {}", reportFileName);
                            System.out.println("檔案名稱 : " + reportFileName);
                            //開始讀取報表
                            try(BufferedReader br = new BufferedReader(new FileReader(reportFile))) {
                                String line;
                                countLine = 0;

                                //只讀取報表前 15行
                                for(int i = 0; i < 15; i++) {
                                    countLine++;
                                    line = br.readLine();
                                    if(line == null) {
                                        logger.info("E-1-2. 檔案 {} 第 {}行之後為空，請檢查檔案", reportFileName, i + 1);
                                        break;
                                    }
                                    //                        System.out.println("字串長度 ： " + line.length());
                                    //                        for (int j = 0; j < line.length(); j++) {
                                    //                            char c = line.charAt(j);
                                    //                            System.out.println("字元: '" + c + "' Unicode: " + (int)c);
                                    //                        }

                                    //讀取到的行有 "日期："
                                    if(line.contains("日期：")){
                                        int startIndex = line.indexOf("日期：");
                                        //Mac 用
//                                    String dateString = line.substring(startIndex + 3, startIndex + 13).replace("/", "");
                                        //Windows 用
                                        String dateString = line.substring(startIndex + 4, startIndex + 14).replace("/", "");

                                        String fileNameDate = reportFileName.substring(reportFileName.lastIndexOf("_") + 1, reportFileName.lastIndexOf("_") + 9);
                                        System.out.println("內文日期 : " + dateString);
                                        System.out.println("檔案名稱的日期 : " + fileNameDate);
                                        logger.info("E-1-2. 內文日期： {}", dateString);
                                        logger.info("E-1-3. 檔案名稱的日期： {}", fileNameDate);

                                        //最後所有正確檔案位置
                                        String updatedFileFolderPath = paramsContent.getUpdatedFileFolderPathURI();
                                        File updatedReportFolder = new File(updatedFileFolderPath + "/" + reportFolderName);
                                        //原檔案 Path
                                        Path sourcePath = reportFile.toPath();
                                        //檔案複製到目標資料夾的 Path
                                        String newFileName = reportFileName.substring(0, reportFileName.lastIndexOf("_") + 1) + dateString + ".txt";
                                        Path targetPath = updatedReportFolder.toPath().resolve(newFileName);
                                        if(!updatedReportFolder.exists()) {
                                            if(updatedReportFolder.mkdirs()) {
                                                logger.info("E-2-1. 建立報表資料夾： {}", reportFolder);
                                            } else {
                                                logger.info("E-2-2. 建立資料夾: {} 失敗！", updatedReportFolder);
                                            }
                                        } else {
                                            logger.info("E-2-3. 報表資料夾： {} 已存在！", reportFolder);
                                        }

                                        boolean isDateMatch = Objects.equals(dateString, fileNameDate);

                                        String updatedPartReportNameDate = reportFileName.substring(reportFileName.indexOf("_") + 1, reportFileName.indexOf("_") + 3) + dateString;

                                        if(!isDateMatch) {
                                            needUpdate++;
                                            System.out.println("日期不符合!!!");
                                            logger.info("E-3. 日期不符合!!!");
                                            needUpdateDate.add(reportFileNameDate);
                                            UpdatedDate.add(updatedPartReportNameDate);
                                        } else {
                                            logger.info("E-3. 日期符合!!!");
                                            originalCorrectCount++;
                                            originalCorrectDate.add(reportFileNameDate);
                                        }

                                        String logMessage = isDateMatch ? "E-4-1. 報表： {} 以複製到指定資料夾！" : "E-4-1. 檔名校正後報表： {} 以複製到指定資料夾！";

                                        try {
                                            Files.copy(sourcePath, targetPath, StandardCopyOption.REPLACE_EXISTING);
                                            dateSet.remove(isDateMatch ? reportFileNameDate : updatedPartReportNameDate);
                                            logger.info(logMessage, targetPath);
                                        } catch(IOException e) {
                                            logger.error("E-4-2. 複製檔案失敗：來源={}，目標={}", sourcePath, targetPath, e);
                                        }

                                        //有找到 "日期："，不需繼續讀完 15行
                                        break;
                                    }
                                }
                            } catch (IOException e) {
                                System.out.println("報表有問題");
                            }
                        }
                    } else {
                        logger.info("E-1. 檔案： {}，不是 .txt 檔", reportFile.getName());
                    }

                    if(countLine >= 15) {
                        System.out.println("讀取 " + countLine + "行後未找到關鍵字，可能為英文報表");
                        logger.info("E-1-2. 讀取第 1個檔案  {} 行後未找到關鍵字，可能為英文報表！", countLine);
                        isEnglishReport = true;
                        //只讀 1個檔案
                        break;
                    }
                }

                System.out.println(dateSet);
                if(dateSet.isEmpty()) {
                    correctMap.put(reportFolderName, reportFolderFiles.length);
                    logger.info("E-5. 需補檔報表日期： {}，不需補檔！", dateSet);
                    continue;
                }
                logger.info("E-5. 需補檔報表日期： {}", dateSet);
            } else if(fileNameList.size() > db2024ExcelMap.get(reportFolderName)) {
                failuredColumnList.add(fileNameList.size());
                failuredColumnList.add(currentCompare2024Count);
                failuredColumnList.add("數量大於 2024年");
                failureMap.put(reportFolderName, failuredColumnList);
                System.out.println(reportFolderName + "需確認報表數量！");
                logger.info("E. 報表資料夾： {} 多於 RIS 系統上數量！請確認報表！", reportFolder);
                continue;
            } else {
                failuredColumnList.add(fileNameList.size());
                failuredColumnList.add(currentCompare2024Count);
                failuredColumnList.add("數量小於 2024年");
                failureMap.put(reportFolderName, failuredColumnList);
                System.out.println(reportFolderName + "需全部重新下載");
                logger.info("E. 報表資料夾： {} 少於 RIS 系統上數量！需全部重新下載！", reportFolder);
                continue;
            }


            //為了不更動到原本的 UpdatedDate，用另一個 Set 變數
            Set<String> updatedCorrect;
            updatedCorrect = UpdatedDate;
            updatedCorrect.removeAll(originalCorrectDate);
            updatedCorrectCount =  updatedCorrect.size();

            //最後需下載日期 (去掉開頭 "數字_" )，並照日期從小到大排序
            Set<String> sortedDates = dateSet.stream()
                    .map(s -> s.substring(s.indexOf("_") + 1)) // 取 "_" 後面的字串
                    .collect(Collectors.toCollection(TreeSet::new)); // TreeSet 自動排序
            Set<String> dateSetSorted = new TreeSet<>(needUpdateDate);

            needUpdatedColumnList.add(totalCount);
            needUpdatedColumnList.add(originalCorrectCount);
            needUpdatedColumnList.add(needUpdate);
            needUpdatedColumnList.add(failureCount);
            needUpdatedColumnList.add(updatedCorrectCount);
            needUpdatedColumnList.add(sortedDates.size());
            needUpdatedColumnList.add(sortedDates);

            if(failureFileName.isEmpty()) {
                needUpdatedColumnList.add("無");
            } else {
                needUpdatedColumnList.add(failureFileName);
            }

            needUpdatedColumnList.add(dateSetSorted);

            if(!isEnglishReport) {
                needUpdatedMap.put(reportFolderName, needUpdatedColumnList);
            }
        }

        //根據與 2024年比較的差異值由大到小排序
        Map<String,List<Object>> sortedByDifferentValueFailureMap =
                failureMap
                        .entrySet()
                        .stream()
                        .sorted( (e1, e2) ->
                                {
                                    Double v1 = (Double) e1.getValue().get(1);
                                    Double v2 = (Double) e2.getValue().get(1);
                                    return v2.compareTo(v1);
                                }
                        )
                        .collect(Collectors
                                .toMap(
                                    Map.Entry::getKey,
                                    Map.Entry::getValue,
                                    (e1, e2) -> e1,
                                    LinkedHashMap::new
                                )
                        );

        //匯出 Excel
        logger.info("F-1-1. 要彙整成 Excel 的 Correct Map： {}", correctMap);
        logger.info("F-1-2. 要彙整成 Excel 的 Need Updated Map： {}", needUpdatedMap);
        logger.info("F-1-3. 要彙整成 Excel 的 Failure Map： {}", sortedByDifferentValueFailureMap);
        String exportPath = paramsContent.getCompareResultExportPath();
        String excelExportPartName = currentFolders.getName();
        System.out.println("資料彙整後的 Excel 存放路徑：" + exportPath);
        logger.info("F-2. 資料彙整後的 Excel 存放路徑： {}", exportPath);
        boolean dataExportResult = TransferMethods.exportXlsx(exportPath, correctMap, needUpdatedMap, sortedByDifferentValueFailureMap, excelExportPartName);

        if(dataExportResult) {
            System.out.println("Excel 匯出成功");
            logger.info("F-3. Excel 匯出成功");
        } else {
            System.out.println("Excel 匯出失敗");
            logger.info("F-3. Excel 匯出失敗");
        }
    }
}