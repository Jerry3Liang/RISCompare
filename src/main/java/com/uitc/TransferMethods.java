package com.uitc;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.*;
import java.util.stream.Collectors;

public class TransferMethods {
    private static final Logger logger = LoggerFactory.getLogger(Main.class);

    public static List<Object> getNeedUpdatedColumnList(
            File reportFolder,
            File[] reportFolderFiles,
            String searchYear,
            int filesCount,
            Set<String> dateSet,
            String updatedFileFolderPath,
            String reportFolderName,
            boolean isEnglishReport
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
                        logger.debug("E-2-1. 檔案名稱 : {} 為空白報表", reportFileName);
                        continue;
                    }

                    logger.info("E-2-1. 檔案名稱： {}", reportFileName);
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
                                logger.info("E-2-2. 檔案 {} 第 {}行之後為空，請檢查檔案", reportFileName, i + 1);
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
//                                String dateString = line.substring(startIndex + 3, startIndex + 13).replace("/", "");
                                //Windows 用
                                String dateString = line.substring(startIndex + 4, startIndex + 14).replace("/", "");

                                String fileNameDate = reportFileName.substring(reportFileName.lastIndexOf("_") + 1, reportFileName.lastIndexOf("_") + 9);
                                System.out.println("內文日期 : " + dateString);
                                System.out.println("檔案名稱的日期 : " + fileNameDate);
                                logger.info("E-2-2. 內文日期： {}", dateString);
                                logger.info("E-2-3. 檔案名稱的日期： {}", fileNameDate);

                                //最後所有正確檔案位置
                                File updatedReportFolder = new File(updatedFileFolderPath + "/" + reportFolderName);
                                //原檔案 Path
                                Path sourcePath = reportFile.toPath();
                                //檔案複製到目標資料夾的 Path
                                String newFileName = reportFileName.substring(0, reportFileName.lastIndexOf("_") + 1) + dateString + ".txt";
                                Path targetPath = updatedReportFolder.toPath().resolve(newFileName);
                                if(!updatedReportFolder.exists()) {
                                    if(updatedReportFolder.mkdirs()) {
                                        logger.info("E-3-1. 建立報表資料夾： {}", reportFolder);
                                    } else {
                                        logger.info("E-3-2. 建立資料夾: {} 失敗！", updatedReportFolder);
                                    }
                                } else {
                                    logger.info("E-3-3. 報表資料夾： {} 已存在！", reportFolder);
                                }

                                boolean isDateMatch = Objects.equals(dateString, fileNameDate);

                                String updatedPartReportNameDate = reportFileName.substring(reportFileName.indexOf("_") + 1, reportFileName.indexOf("_") + 3) + dateString;

                                if(!isDateMatch) {
                                    needUpdate++;
                                    System.out.println("日期不符合!!!");
                                    logger.info("E-4. 日期不符合!!!");
                                    needUpdateDate.add(reportFileNameDate);
                                    UpdatedDate.add(updatedPartReportNameDate);
                                } else {
                                    logger.info("E-4. 日期符合!!!");
                                    originalCorrectCount++;
                                    originalCorrectDate.add(reportFileNameDate);
                                }

                                String logMessage = isDateMatch ? "E-5-1. 報表： {} 以複製到指定資料夾！" : "E-5-1. 檔名校正後報表： {} 以複製到指定資料夾！";

                                try {
                                    Files.copy(sourcePath, targetPath, StandardCopyOption.REPLACE_EXISTING);
                                    dateSet.remove(isDateMatch ? reportFileNameDate : updatedPartReportNameDate);
                                    logger.info(logMessage, targetPath);
                                } catch(IOException e) {
                                    logger.error("E-5-2. 複製檔案失敗：來源={}，目標={}", sourcePath, targetPath, e);
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
                logger.info("E-2-1. 檔案： {}，不是 .txt 檔", reportFile.getName());
            }

            if(countLine >= 15) {
                System.out.println("讀取 " + countLine + "行後未找到關鍵字，可能為英文報表");
                logger.info("E-2-1. 讀取第 1個檔案  {} 行後未找到關鍵字，可能為英文報表！", countLine);
                isEnglishReport = true;
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
