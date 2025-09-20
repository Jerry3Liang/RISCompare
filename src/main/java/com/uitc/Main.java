package com.uitc;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.stream.Collectors;

import static com.uitc.DateConverter.convertDates;
import static com.uitc.ReadExcelMethods.getParamsContent;
import static com.uitc.ReadExcelMethods.getReportCountsAndExpYearForYear;
import static com.uitc.TransferMethods.*;
import static com.uitc.WriteExcelMethods.exportIntegrationXlsx;
import static com.uitc.WriteExcelMethods.exportLackReportXlsx;

public class Main {

    private static final Logger logger = LoggerFactory.getLogger(Main.class);

    public static void main(String[] args) throws IOException, InterruptedException {
        //Mac 用
//        final String PARAMS_PATH = "/Users/liangchejui/Desktop/程式語言相關/RIS 比對小程式資料/參數檔/params_for_RIS.xlsx";
        //Windows 用
        final String PARAMS_PATH = "C:/比對小程式/Config/params_for_RIS.xlsx";
        logger.info("A. 開始讀取參數檔，參數檔 Path： {}", PARAMS_PATH);

        //讀取參數檔
        ParamsContent paramsContent = getParamsContent(PARAMS_PATH);

        //要比對的年份
        String searchYear = paramsContent.getReportInYear();
        logger.info("B. 比對報表年份： {}", searchYear);

        //當前年份
        Date date = new Date();
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        int year = calendar.get(Calendar.YEAR);
        logger.info("C. 當前年份： {}", year);

        //取得報表代號對應總份數的 Map
        String yearExcelPath = paramsContent.getYearExcelPath();
        Map<String, List<Double>> yearExcelMap = getReportCountsAndExpYearForYear(yearExcelPath);
        logger.info("D. {}年報表代號對應總份數 Excel 所在檔案路徑： {}", searchYear, yearExcelPath);
//        System.out.println("yearExcelMap : " + yearExcelMap);

        String searchFolderPathUrl = paramsContent.getSearchFolderPathURI();
        logger.info("E. 報表所在資料夾路徑： {}", searchFolderPathUrl);
        File currentFolders = new File(searchFolderPathUrl);

        //最後匯出 Excel 需要的資料
        //1. 數量正確且不需校正
        Map<String, Integer> correctMap = new HashMap<>();

        //2. 數量正確但需要校正
        Map<String, List<Object>> needUpdatedMap = new HashMap<>();

        //3. 數量不正確，比對後需重新下載
        Map<String, List<Object>> failureMap = new HashMap<>();

        //4. 沒有下載到報表
        List<String> noDownloadReportList = new ArrayList<>();

        //5. 過期報表，不需轉入 RIS
        List<String> expYearReportIdList = new ArrayList<>();

        //最後所有正確檔案位置
        String updatedFileFolderPath = paramsContent.getUpdatedFileFolderPathURI();

        //以 yearExcelMap 為主
        List<String> keys = new ArrayList<>(yearExcelMap.keySet());
        Collections.sort(keys);
        for(String key : keys) {
            if((year - Integer.parseInt(searchYear)) <= yearExcelMap.get(key).get(1)) {
                //紀錄以比對過
                boolean isCompared = false;

                File[] reportFolders = currentFolders.listFiles(File::isDirectory);
                assert reportFolders != null;

                //遍歷所有報表代號的資料夾
                for (File reportFolder : reportFolders) {
                    //取得報表資料夾的名稱
                    String reportFolderName = reportFolder.getName();

                    if(key.equals(reportFolderName)) {
                        isCompared = true;
                        System.out.println("報表資料夾 : " + reportFolderName);
                        logger.info("E-1-1. 報表資料夾： {}", reportFolderName);

                        //判斷是否為英文報表
                        boolean isEnglishReport = false;

                        //2. 的欄位資料
                        List<Object> needUpdatedColumnList;
                        //建立全部檔案名稱日期 Set
                        Set<String> dateSet;

                        //第 3. 的欄位資料
                        List<Object> failuredColumnList = new ArrayList<>();

                        //打開報表代號的資料夾並過濾掉不是 file 及隱藏檔案
                        File[] reportFolderFiles = reportFolder.listFiles(file -> file.isFile() && !file.isHidden());
                        assert reportFolderFiles != null;

                        //建立檔案名稱的 List
                        List<String> fileNameList = getFileNameList(reportFolderFiles, searchYear);

                        int filesCount = fileNameList.size();

                        //OnDemand 查詢到的份數與實際下載到資料夾的份數差異
                        double currentCompareYearExcelCount = Math.abs(yearExcelMap.get(reportFolderName).get(0) - filesCount);

                        if(filesCount == yearExcelMap.get(reportFolderName).get(0)) {
                            logger.info("E-2. 報表資料夾： {} 與 OnDemand 系統上數量一樣！", reportFolderName);
                            //遍歷報表代號的資料夾裡所有檔案並將檔案名稱的日期存取到 dateSet
                            dateSet = getDateSet(reportFolderFiles, searchYear);
                            needUpdatedColumnList = getNeedUpdatedColumnList(reportFolder, reportFolderFiles, searchYear, filesCount, dateSet, updatedFileFolderPath, reportFolderName, isEnglishReport);

                            System.out.println(dateSet);
                            if(dateSet.isEmpty()) {
                                correctMap.put(reportFolderName, filesCount);
                                logger.info("E-6. 需補檔報表日期： {}，不需補檔！", dateSet);
                                continue;
                            }
                            logger.info("E-6. 需補檔報表日期： {}", dateSet);
                        } else if(filesCount > yearExcelMap.get(reportFolderName).get(0)) {
                            failuredColumnList.add(fileNameList.size());
                            failuredColumnList.add(currentCompareYearExcelCount);
                            failuredColumnList.add("數量大於 OnDemand 系統上 " + searchYear + "年的數量");
                            failureMap.put(reportFolderName, failuredColumnList);
                            System.out.println(reportFolderName + "需確認報表數量！");
                            logger.info("E-2. 報表資料夾： {} 多於 OnDemand 系統上數量！請確認報表！", reportFolderName);
                            continue;
                        } else {
                            failuredColumnList.add(fileNameList.size());
                            failuredColumnList.add(currentCompareYearExcelCount);
                            failuredColumnList.add("數量小於 OnDemand 系統上 " + searchYear + "年的數量");
                            failureMap.put(reportFolderName, failuredColumnList);
                            System.out.println(reportFolderName + "需全部重新下載");
                            logger.info("E-2. 報表資料夾： {} 少於 OnDemand 系統上數量！需全部重新下載！", reportFolderName);
                            continue;
                        }

                        if(!isEnglishReport) {
                            needUpdatedMap.put(reportFolderName, needUpdatedColumnList);
                        }
                    }
                }

                if(!isCompared) {
                    noDownloadReportList.add(key);
                    logger.info("E-1-2. 沒有名稱為 {} 的報表資料夾！", key);
                }
            } else {
                expYearReportIdList.add(key);
                logger.info("E. 報表： {}，已過期不需轉入 RIS 系統", key);
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

        //匯出統整 Excel
        logger.info("F. 開始匯出統整 Excel");
        logger.info("F-1-1. 要彙整成統整 Excel 的 Correct Map： {}", correctMap);
        logger.info("F-1-2. 要彙整成統整 Excel 的 Need Updated Map： {}", needUpdatedMap);
        logger.info("F-1-3. 要彙整成統整 Excel 的 Failure Map： {}", sortedByDifferentValueFailureMap);
        logger.info("F-1-4. 要彙整成統整 Excel 的 No Download Report List： {}", noDownloadReportList);
        logger.info("F-1-5. 要彙整成統整 Excel 的 Exp. Year Report List： {}", expYearReportIdList);
        String exportPath = paramsContent.getCompareResultExportPath();
        String excelExportPartName = currentFolders.getName();
        System.out.println("資料彙整後的 Excel 存放路徑：" + exportPath);
        logger.info("F-2. 資料彙整後的 Excel 存放路徑： {}", exportPath);
        boolean integrationDataExportResult = exportIntegrationXlsx(
                exportPath,
                correctMap, needUpdatedMap, sortedByDifferentValueFailureMap, noDownloadReportList, expYearReportIdList,
                excelExportPartName
        );

        if(integrationDataExportResult) {
            System.out.println("統整 Excel 匯出成功");
            logger.info("F-3. 統整 Excel 匯出成功");
        } else {
            System.out.println("統整 Excel 匯出失敗");
            logger.info("F-3. 統整 Excel 匯出失敗");
        }

        //開始處理報表數量與 OnDemand 上數量不同的問題
        if(!sortedByDifferentValueFailureMap.isEmpty()) {
            logger.info("G. 開始處理報表數量與 OnDemand 上數量不同的問題");
            //最後匯出 Excel 需要的資料
            //已存在檔案
            Map<String, List<Object>> excel2ndExistMap = new HashMap<>();
            //原本就缺檔案
            Map<String, List<Object>> excel2ndNotExistMap = new HashMap<>();

            List<String> sortedByDifferentValueFailureMapKeys = new ArrayList<>(sortedByDifferentValueFailureMap.keySet());
            Collections.sort(sortedByDifferentValueFailureMapKeys);

            for(String failureMapKey : sortedByDifferentValueFailureMapKeys) {

                File[] failureReportFolders = currentFolders.listFiles(File::isDirectory);
                assert failureReportFolders != null;

                //遍歷所有有問題報表代號的資料夾
                for (File failureReportFolder : failureReportFolders) {
                    //取得報表資料夾的名稱
                    String failureReportFolderName = failureReportFolder.getName();

                    if(failureMapKey.equals(failureReportFolderName)) {

                        System.out.println("有缺報表資料夾 : " + failureReportFolderName);
                        logger.info("G-1 有缺報表資料夾： {}", failureReportFolderName);

                        //判斷是否為英文報表
                        boolean isEnglishReport = false;

                        //Excel 欄位資料
                        //已存在
                        List<Object> excel2ndExistData;
                        //建立全部檔案名稱日期 Set
                        Set<String> dateSet;
                        //不存在
                        List<Object> excel2ndNotExistData = new ArrayList<>();

                        //先填欄位1
                        excel2ndNotExistData.add(sortedByDifferentValueFailureMap.get(failureMapKey).get(1));

                        //打開報表代號的資料夾並過濾掉不是 file 及隱藏檔案
                        File[] failureReportFolderFiles = failureReportFolder.listFiles(file -> file.isFile() && !file.isHidden());
                        assert failureReportFolderFiles != null;

                        //建立檔案名稱的 List
                        List<String> fileNameList = getFileNameList(failureReportFolderFiles, searchYear);
                        int countFiles = fileNameList.size();
                        //遍歷報表代號的資料夾裡所有檔案並將檔案名稱的日期存取到 dateSet
                        dateSet = getDateSet(failureReportFolderFiles, searchYear);
                        excel2ndExistData = getNeedUpdatedColumnList(failureReportFolder, failureReportFolderFiles, searchYear, countFiles, dateSet, updatedFileFolderPath, failureReportFolderName, isEnglishReport);

                        System.out.println(dateSet);
                        if(dateSet.isEmpty()) {
                            logger.info("G-2. 已存在但需補檔報表日期： {}，不需補檔！", dateSet);
                        } else {
                            logger.info("G-2. 已存在但需補檔報表日期： {}", dateSet);
                        }

                        if(!isEnglishReport) {
                            excel2ndExistMap.put(failureReportFolderName, excel2ndExistData);
                        }

                        //開始處理不存在的檔案
                        logger.info("G-3. 開始處理缺的報表數量問題");
                        //從 OnDemand 傳回來的 String
                        String onDemandString;
                        // 1. 啟動 AutoIt 腳本打開 OnDemand 搜尋有缺的報表代號
                        //先登入
                        logger.info("G-3-1. 開始啟動 OnDemand 並登入");
                        ProcessBuilder loginPb = new ProcessBuilder(
                                "C:\\Program Files (x86)\\AutoIt3\\AutoIt3.exe",
                                "C:\\autoIt3Script\\autoOndemandLogin.au3"
                        );
                        loginPb.redirectErrorStream(true);
                        Process loginProc = loginPb.start();
                        loginProc.waitFor();  //等待登入完成
                        logger.info("G-3-2. OnDemand 登入成功");
                        logger.info("G-3-3. OnDemand 開始查詢報表");
                        ProcessBuilder searchReportPb = new ProcessBuilder(
                                "C:\\Program Files (x86)\\AutoIt3\\AutoIt3.exe",
                                "C:\\autoIt3Script\\autoOndemandSearchReportIdMonthCounts.au3",
                                failureReportFolderName,
                                searchYear
                        );
                        searchReportPb.redirectErrorStream(true);
                        Process process = searchReportPb.start();
                        //讀取 AutoIt 輸出的結果
                        try (BufferedReader br = new BufferedReader(
                                new InputStreamReader(process.getInputStream(), StandardCharsets.UTF_8))
                        ) {
                                StringBuilder sb = new StringBuilder();
                                String line;
                                while ((line = br.readLine()) != null) {
                                    sb.append(line);
                                }

                                onDemandString = sb.toString();
                            }

                            process.waitFor();
                            logger.info("G-3-4. 報表查詢成功");

                            List<String> dateList = convertDates(onDemandString);
                            logger.info("G-3-5 開始比對缺的報表");
                            dateList.removeAll(fileNameList);
                            logger.info("G-3-5 缺的報表比對完成，所缺檔案： {}", dateList);
                            //最後需下載日期 (去掉開頭 "數字_" )，並照日期從小到大排序
                            Set<String> sortedDates = dateList.stream()
                                    .map(s -> s.substring(s.indexOf("_") + 1)) // 取 "_" 後面的字串
                                    .collect(Collectors.toCollection(TreeSet::new)); // TreeSet 自動排序
                            //填欄位2
                            excel2ndNotExistData.add(sortedDates);

                            //填欄位3
                            excel2ndNotExistData.add(dateList);

                            excel2ndNotExistMap.put(failureMapKey, excel2ndNotExistData);
                            System.out.println(excel2ndNotExistMap);
                    }
                }
            }

            //匯出有缺報表 Excel
            logger.info("H. 開始匯出有缺報表 Excel");
            logger.info("H-1-1. 要彙整成有缺報表 Excel 的 Exist Map： {}", excel2ndExistMap);
            logger.info("H-1-2. 要彙整成有缺報表 Excel 的 Not Exist Map： {}", excel2ndNotExistMap);
            logger.info("H-2. 有缺報表的 Excel 存放路徑： {}", exportPath);
            boolean LackReportDataExportResult = exportLackReportXlsx(
                    exportPath,
                    excel2ndExistMap, excel2ndNotExistMap,
                    excelExportPartName
            );

            if(LackReportDataExportResult) {
                System.out.println("有缺報表 Excel 匯出成功");
                logger.info("H-3. 有缺報表 Excel 匯出成功");
            } else {
                System.out.println("有缺報表 Excel 匯出失敗");
                logger.info("H-3. 有缺報表 Excel 匯出失敗");
            }
        }
    }
}