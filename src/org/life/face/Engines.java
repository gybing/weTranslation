package org.life.face;

import org.life.Core.Excel.ExcelProcessor;
import org.life.Core.Excel.Interface.Excel;
import org.life.Core.FileObjectFactory;
import org.life.Core.FileReaders;
import org.life.Core.TypeFlag.AbstractFlag.Flag;
import org.life.Core.TypeFlag.HSSFFlag;
import org.life.Core.TypeFlag.XSSFFlag;
import org.life.Core.TypeFlag.XWPFFlag;
import org.life.Core.Word.Interface.Word;
import org.life.Core.Word.WordProcessor;

import java.util.Map;

public final class Engines {
    private static Excel excelHandles;
    private static Word wordHandles;
    private static String excelPath = "";
    private static String wordPath = "";

    private Engines() {}
    public static void initExcel(String path, int sheetNum, int metadataRowNum) {
        if(excelPath.equals(path))return;
        else excelPath = path;

        Flag flag = FileReaders.getFileFlag(excelPath);
        if(flag instanceof HSSFFlag) {
            excelHandles = new ExcelProcessor.Builder(FileObjectFactory
                    .getWorkBook((HSSFFlag) flag))
                    .setSheetIndex(sheetNum)
                    .setMetaDataRow(metadataRowNum)
                    .build();
        } else {
            excelHandles = new ExcelProcessor.Builder(FileObjectFactory
                    .getWorkBook((XSSFFlag) flag))
                    .setSheetIndex(sheetNum)
                    .setMetaDataRow(metadataRowNum)
                    .build();
        }
    }

    public static void initWordInReadModel(String path, Map<String, Integer> directionMap,
                               Map<String, Integer> metaDataAndLengthMap, String[] symbols) {
        if(wordPath.equals(path))return;
        else wordPath = path;

        Flag flag = FileReaders.getFileFlag(wordPath);
        wordHandles = new WordProcessor.Builder(FileObjectFactory
                .getDocument((XWPFFlag) flag))
                .setTrimSet(symbols)
                .setDataDirectionMap(directionMap)
                .setMetaDataAndLengthMap(metaDataAndLengthMap)
                .build();
    }

    public static void initWordInWriteModel(String path, Map<String, Integer> directionMap, String[] symbols) {
        if(wordPath.equals(path))return;
        else wordPath = path;

        Flag flag = FileReaders.getFileFlag(wordPath);
        wordHandles = new WordProcessor.Builder(FileObjectFactory
                .getDocument((XWPFFlag) flag))
                .setDataDirectionMap(directionMap)
                .setTrimSet(symbols)
                .build();
    }

    public static Excel getExcelHandles()
    {
        return excelHandles;
    }

    public static Word getWordHandles()
    {
        return wordHandles;
    }

    public static boolean isSameExcel(String path)
    {
        return excelPath.equals(path);
    }

    public static boolean isSameWord(String path)
    {
        return wordPath.equals(path);
    }
}
