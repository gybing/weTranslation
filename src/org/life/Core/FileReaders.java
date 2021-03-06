package org.life.Core;

import org.life.Core.TypeFlag.AbstractFlag.Flag;
import org.life.Core.TypeFlag.HSSFFlag;
import org.life.Core.TypeFlag.XSSFFlag;
import org.life.Core.TypeFlag.XWPFFlag;
import org.life.Exception.FileNotExistException;
import org.life.Exception.FlagException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

/**
 * 一个返回 File InputStream 的类，同时返回该文件的类型标记
 */
public final class FileReaders {
    private static final String XLS = "xls";
    private static final String XLSX = "xlsx";
    private static final String DOCX = "docx";

    /**
     * 检查文后缀是否符合要求
     * @param path 文件路径
     * @return 符合要求返回true, 否则返回false
     */
    public static boolean checkSuffix(String path)
    {
        String[] tmpArr = path.split("\\.");
        switch (tmpArr[tmpArr.length - 1])
        {
            case XLS:
            case XLSX:
            case DOCX: return true;
            default: return false;
        }
    }

    /**
     * 获取文件类型标记
     * @param path 文件路径
     * @return 返回一个抽象文件类型标记
     */
    public static Flag getFileFlag(String path)
    {
        File fileChecker = new File(path);
        if(! fileChecker.exists())throw new FileNotExistException("file: " + path + " is not exist.");

        FileInputStream fileStream = null;
        try
        {
            fileStream = new FileInputStream(fileChecker);
        }
        catch (FileNotFoundException e)
        {
            e.printStackTrace();
            System.exit(1);
        }

        String[] tmpArr = fileChecker.getName().split("\\.");
        switch (tmpArr[tmpArr.length - 1])
        {
            case XLS: return new HSSFFlag(fileStream);
            case XLSX: return new XSSFFlag(fileStream);
            case DOCX: return new XWPFFlag(fileStream);
            default: throw new FlagException(String
                    .format("illegal flag: %s is unknown file type.", tmpArr[tmpArr.length - 1]));
        }
    }
}
