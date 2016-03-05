package org.life.Core;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.life.Core.TypeFlag.HSSFFlag;
import org.life.Core.TypeFlag.XSSFFlag;
import org.life.Core.TypeFlag.XWPFFlag;

import java.io.IOException;

/**
 * 创建各类型的文件操作对象的工厂
 */
public final class FileObjectFactory {
    private FileObjectFactory() {}

    public static HSSFWorkbook getWorkBook(HSSFFlag hssfFlag)
    {
        try {
            return new HSSFWorkbook(hssfFlag.getStream());
        }
        catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException(String
                    .format("%s: get HSSFWorkbook failure (Method: getWorkBook).", FileObjectFactory.class.getName()));
        }
    }

    public static XSSFWorkbook getWorkBook(XSSFFlag xssfFlag)
    {
        try {
            return new XSSFWorkbook(xssfFlag.getStream());
        }
        catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException(String
                    .format("%s: get XSSFWorkbook failure (Method: getWorkBook).", FileObjectFactory.class.getName()));
        }
    }

    public static XWPFDocument getDocument(XWPFFlag xwpfFlag)
    {
        try {
            return new XWPFDocument(xwpfFlag.getStream());
        }
        catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException(String
                    .format("%s: get XWPFDocument failure (Method: getDocument).", FileObjectFactory.class.getName()));
        }
    }
}
