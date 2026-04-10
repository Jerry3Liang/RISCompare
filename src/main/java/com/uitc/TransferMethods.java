package com.uitc;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.*;
import java.util.stream.Collectors;

import static com.uitc.DateConverter.extractDatesFromFile;

public class TransferMethods {
    private static final Logger logger = LoggerFactory.getLogger(TransferMethods.class);

    private static final String[] oldReportId = {"ABLB191A", "ABLB198A", "ABLB348B", "ABLB902A", "ABLB902B", "ABLB902C",
                                                 "ABLB902D", "ABLB902E", "ABLR299A", "ACMB041B", "ACQB104A", "ACQB104C",
                                                 "ACQB104D", "ACQB104E", "ACQB104F", "ACQB104G", "ACQB104Z", "AISB038C",
                                                 "AISB219A", "AISB342B", "AISB426A", "CMPB008A", "CMPB286A", "CMPB286B",
                                                 "CMPB286C", "CMPB286E", "CMPB774A", "CMPB832D", "CMPB832G", "CMPB832H"};
    private static final String[] newReportId = {"ABLB191A", "ABLB198A", "ABLB348B", "ABLB902A", "ABLB902B", "ABLB902C",
                                                 "ABLB902D", "ABLB902E", "ABLR299A", "ACMB041B", "ACQB104A", "ACQB104C",
                                                 "ACQB104D", "ACQB104E", "ACQB104F", "ACQB104G", "ACQB104Z", "AISB038C",
                                                 "AISB219A", "AISB342B", "AISB426A", "CMPB008A", "CMPB286A", "CMPB286B",
                                                 "CMPB286C", "CMPB286E", "CMPB774A", "CMPB832D", "CMPB832G", "CMPB832H"};
    static String jsonConfigPath = "/json/specialReportDateFormatter.json";

    public static List<Object> getNeedUpdatedColumnList(
            File reportFolder,
            File[] reportFolderFiles,
            String searchYear,
            int filesCount,
            Set<String> dateSet,
            String updatedFileFolderPath,
            String reportFolderName,
            boolean isEArea
    ) {
        List<Object> needUpdatedColumnList = new ArrayList<>();

        int originalCorrectCount = 0;
        int needUpdate = 0;
        int updatedCorrectCount;
        int failureCount = 0;

        //建立錯誤報表檔案名稱 List
        List<String> failureFileName = new ArrayList<>();
        //建立日期原本就正確 Set
        Set<String> originalCorrectDate = new HashSet<>();
        //建立需校正名單
        Set<String> needUpdateDate = new HashSet<>();
        //建立校正完名稱 Set，儲存改完名的日期
        Set<String> UpdatedDate = new HashSet<>();

        //計算讀取報表內容行數
        int countLine = 0;

        //遍歷報表代號的資料夾裡所有檔案
        for (File reportFile : reportFolderFiles) {
            if(isValidFile(reportFile)) {
                String reportFileName = reportFile.getName();
                String reportFileNameYear = reportFileName.substring(reportFileName.lastIndexOf("_") + 1, reportFileName.lastIndexOf("_") + 5);
                if(Objects.equals(searchYear, reportFileNameYear)) {
                    String reportFileNameDate = reportFileName.substring(reportFileName.indexOf("_") + 1, reportFileName.lastIndexOf("_") + 9);

                    //Windows 電腦，檔案如果為空白，容量為 0
                    if(reportFile.length() == 0) {
                        failureCount++;
                        failureFileName.add(reportFileName);
                        String logMessage1 = isEArea ? "E-2-1. 檔案名稱 : {} 為空白報表" : "G-2-1. 檔案名稱 : {} 為空白報表";
                        logger.debug(logMessage1, reportFileName);
                        continue;
                    }

                    String logMessage2 = isEArea ? "E-2-1. 檔案名稱： {}" : "G-2-1. 檔案名稱： {}";
                    logger.info(logMessage2, reportFileName);
                    System.out.println("檔案名稱 : " + reportFileName);

                    //開始讀取報表
                    try(BufferedReader br = new BufferedReader(new FileReader(reportFile))) {
                        countLine = 0;

                        ConfigLoader loader = new ConfigLoader(jsonConfigPath);
                        //只讀取報表前 15行
                        for(int i = 0; i < 15; i++) {
                            countLine++;
                            String line = br.readLine();
                            if(line == null) {
                                String logMessage3 = isEArea ? "E-2-2. 檔案 {} 第 {}行之後為空，請檢查檔案" : "G-2-2. 檔案 {} 第 {}行之後為空，請檢查檔案";
                                logger.info(logMessage3, reportFileName, i + 1);
                                break;
                            }
                            //                        System.out.println("字串長度 ： " + line.length());
                            //                        for (int j = 0; j < line.length(); j++) {
                            //                            char c = line.charAt(j);
                            //                            System.out.println("字元: '" + c + "' Unicode: " + (int)c);
                            //                        }

                            String reportContentDate = extractDatesFromFile(
                                                        line,
                                                        loader,
                                                        reportFileName.substring(0, reportFileName.indexOf("_"))
                            );

                            if(!reportContentDate.contains("無法辨識格式")){
                                String fileNameDate = reportFileName.substring(reportFileName.lastIndexOf("_") + 1, reportFileName.lastIndexOf("_") + 9);
                                System.out.println("內文日期 : " + reportContentDate);
                                System.out.println("檔案名稱的日期 : " + fileNameDate);
                                String logMessage4 = isEArea ? "E-2-2. 內文日期： {}" : "G-2-2. 內文日期： {}";
                                logger.info(logMessage4, reportContentDate);
                                String logMessage5 = isEArea ? "E-2-3. 檔案名稱的日期： {}" : "G-2-3. 檔案名稱的日期： {}";
                                logger.info(logMessage5, fileNameDate);

                                //先確認 ReportId 是否為 oldReportId
                                String checkedReportIdName = replaceIfExists(oldReportId, newReportId, reportFolderName);
                                System.out.println("checkedReportIdName : " + checkedReportIdName);

                                //最後所有正確檔案位置
                                File updatedReportFolder = new File(updatedFileFolderPath + "/" + reportFolderName);
                                //原檔案 Path
                                Path sourcePath = reportFile.toPath();
                                //檔案複製到目標資料夾的 Path
                                String newFileName = checkedReportIdName + reportFileName.substring(reportFileName.indexOf("_"), reportFileName.lastIndexOf("_") + 1) + reportContentDate + ".txt";
                                Path targetPath = updatedReportFolder.toPath().resolve(newFileName);
                                System.out.println("newFileName : " + newFileName);

                                if(!updatedReportFolder.exists()) {
                                    if(updatedReportFolder.mkdirs()) {
                                        String logMessage6 = isEArea ? "E-3-1. 建立報表資料夾： {}" : "G-3-1. 建立報表資料夾： {}";
                                        logger.info(logMessage6, reportFolder);
                                    } else {
                                        String logMessage7 = isEArea ? "E-3-2. 建立資料夾: {} 失敗！" : "G-3-2. 建立資料夾: {} 失敗！";
                                        logger.info(logMessage7, updatedReportFolder);
                                    }
                                } else {
                                    String logMessage8 = isEArea ? "E-3-3. 報表資料夾： {} 已存在！" : "G-3-3. 報表資料夾： {} 已存在！";
                                    logger.info(logMessage8, reportFolder);
                                }

                                boolean isDateMatch = Objects.equals(reportContentDate, fileNameDate);

                                String updatedPartReportNameDate = reportFileName.substring(reportFileName.indexOf("_") + 1, reportFileName.indexOf("_") + 3) + reportContentDate;

                                if(!isDateMatch) {
                                    needUpdate++;
                                    System.out.println("日期不符合!!!");
                                    String logMessage9 = isEArea ? "E-4. 日期不符合!!!" : "G-4. 日期不符合!!!";
                                    logger.info(logMessage9);
                                    needUpdateDate.add(reportFileNameDate);
                                    UpdatedDate.add(updatedPartReportNameDate);
                                } else {
                                    String logMessage10 = isEArea ? "E-4. 日期符合!!!" : "G-4. 日期符合!!!";
                                    logger.info(logMessage10);
                                    originalCorrectCount++;
                                    originalCorrectDate.add(reportFileNameDate);
                                }

                                String logMessage11 = isEArea ? "E-5-1. 報表： {} 以複製到指定資料夾！" : "G-5-1. 報表： {} 以複製到指定資料夾！";
                                String logMessage12 = isEArea ? "E-5-1. 檔名校正後報表： {} 以複製到指定資料夾！" : "G-5-1. 檔名校正後報表： {} 以複製到指定資料夾！";

                                String logMessage = isDateMatch ? logMessage11 : logMessage12;

                                try {
                                    Files.copy(sourcePath, targetPath, StandardCopyOption.REPLACE_EXISTING);
                                    dateSet.remove(isDateMatch ? reportFileNameDate : updatedPartReportNameDate);
                                    logger.info(logMessage, targetPath);
                                } catch(IOException e) {
                                    String logMessage13 = isEArea ? "E-5-2. 複製檔案失敗：來源={}，目標={}" : "G-5-2. 複製檔案失敗：來源={}，目標={}";
                                    logger.error(logMessage13, sourcePath, targetPath, e);
                                }

                                //有找到 "日期："，不需繼續讀完 15行
                                break;
                            }
                        }
                    } catch (IOException e) {
                        System.out.println("報表有問題");
                    }
                } else {
                    String logMessage14 = isEArea ? "E-2-1. 要比對的年份： {} 與檔案名稱的年份： {}，不一致！" : "G-2-1. 要比對的年份： {} 與檔案名稱的年份： {}，不一致！";
                    logger.info(logMessage14, searchYear, reportFileNameYear);
                }
            } else {
                String logMessage15 = isEArea ? "E-2-1. 檔案： {}，不是 .txt 檔" : "G-2-1. 檔案： {}，不是 .txt 檔";
                logger.info(logMessage15, reportFile.getName());
            }

            if(countLine >= 15) {
                System.out.println("讀取 " + countLine + "行後未找到關鍵字");
                String logMessage16 = isEArea ? "E-2-1. 讀取第 1個檔案  {} 行後未找到關鍵字！" : "G-2-1. 讀取第 1個檔案  {} 行後未找到關鍵字！";
                logger.info(logMessage16, countLine);
                //只讀 1個檔案
                break;
            }
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

        needUpdatedColumnList.add(filesCount);
        needUpdatedColumnList.add(originalCorrectCount);
        needUpdatedColumnList.add(needUpdate);
        needUpdatedColumnList.add(failureCount);
        needUpdatedColumnList.add(updatedCorrectCount);
        needUpdatedColumnList.add(sortedDates.size());
        needUpdatedColumnList.add(sortedDates);

        int needReDownloadFileCounts = countMatchingDates(sortedDates, dateSetSorted);
        System.out.println(needReDownloadFileCounts);
        needUpdatedColumnList.add(needReDownloadFileCounts);

        if(failureFileName.isEmpty()) {
            needUpdatedColumnList.add("無");
        } else {
            needUpdatedColumnList.add(failureFileName);
        }

        needUpdatedColumnList.add(dateSetSorted);

        return needUpdatedColumnList;
    }

    public static List<String> getFileNameList(File[] reportFolderFiles, String searchYear){
        List<String> fileNameList = new ArrayList<>();

        //先遍歷所有檔案名稱，日期年的位置與要比對的年份一致就存入 fileNameList
        for(File reportFile : reportFolderFiles) {
            String reportFileName = reportFile.getName();
            String reportFileNameDate = reportFileName.substring(reportFileName.indexOf("_") + 1, reportFileName.lastIndexOf("_") + 9);
            String reportFileNameYear = reportFileNameDate.substring(reportFileNameDate.indexOf("_") + 1, reportFileNameDate.indexOf("_") + 5);
            if(Objects.equals(searchYear, reportFileNameYear)) {
                fileNameList.add(reportFileNameDate);
            }
        }

        return fileNameList;
    }

    public static Set<String> getDateSet(File[] reportFolderFiles, String searchYear) {
        //建立全部檔案名稱日期 Set
        Set<String> dateSet = new HashSet<>();

        //遍歷報表代號的資料夾裡所有檔案並將檔案名稱的日期存取到 dateSet
        for(File reportFile : reportFolderFiles) {
            String reportFileName = reportFile.getName();
            String reportFileNameYear = reportFileName.substring(reportFileName.lastIndexOf("_") + 1, reportFileName.lastIndexOf("_") + 5);
            if(Objects.equals(searchYear, reportFileNameYear)) {
                String fileNameDate = reportFileName.substring(reportFileName.indexOf("_") + 1, reportFileName.lastIndexOf("_") + 9);
                dateSet.add(fileNameDate);
            }
        }

        return dateSet;
    }

    private static String replaceIfExists(String[] oldReportId, String[] newReportId, String target) {
        int index = Arrays.asList(oldReportId).indexOf(target);
        if (index != -1) {
            return newReportId[index];
        }

        return target; //找不到就回傳原值
    }

    public static int countMatchingDates(Set<String> finalNeedReDownloadDate, Set<String> needUpdatePartDateName) {
        return (int) needUpdatePartDateName.stream()
                .map(s -> s.split("_"))
                .filter(parts -> parts.length == 2 && finalNeedReDownloadDate.contains(parts[1]))
                .count();
    }

    /**
     * 檢查檔案名稱是否以 ".txt" 或 ".TXT" 結尾 (符合 Excel 檔的副檔名)，及過濾掉檔名開頭是 ._ 或 ~$ 或 .DS。
     * @param file：所有讀取到的檔案。
     * @return boolean: true (是 txt 檔) 或 false (不是 txt 檔)。
     */
    public static boolean isValidFile(File file) {
        //檢查檔案名稱的副檔名是否為 ".txt" 或 ".TXT"。
        if (!(file.getName().endsWith(".txt") || file.getName().endsWith(".TXT"))) {
            return false;
        }

        // 檢查檔案名稱是否以 "._" 或 "~$" 或 ".DS"開頭。
        return !file.getName().startsWith("._") && !file.getName().startsWith("~$") && !file.getName().startsWith(".DS");
    }
}
