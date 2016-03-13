package org.life.Test;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.life.Core.Excel.ExcelProcessor;
import org.life.Core.Excel.Interface.Excel;
import org.life.Core.FileObjectFactory;
import org.life.Core.FileReaders;
import org.life.Core.TypeFlag.AbstractFlag.Flag;
import org.life.Core.TypeFlag.XSSFFlag;

import java.util.Arrays;

public class Test2 {

    public static void main(String[] args)
    {
        Flag flag = FileReaders.getFileFlag("E:\\12.xlsx");
        XSSFWorkbook wb = FileObjectFactory.getWorkBook((XSSFFlag) flag);

        ExcelProcessor.Builder builder = new ExcelProcessor.Builder(wb);
        builder.setSheetIndex(0);
        builder.setMetaDataRow(2);
        Excel process = builder.build();

        System.out.println(Arrays.toString(process.getMetaData().toArray()));
    }
}
