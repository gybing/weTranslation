package org.life.Core.Excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.life.Core.Excel.Interface.ExcelProcessor;
import org.life.Core.FileObjectFactory;
import org.life.Core.FileReaders;
import org.life.Core.TypeFlag.AbstractFlag.Flag;
import org.life.Core.TypeFlag.XSSFFlag;

import java.util.Arrays;

public class Test {

    public static void main(String[] args)
    {
        Flag flag = FileReaders.getFileFlag("E:\\12.xlsx");
        XSSFWorkbook wb = FileObjectFactory.getWorkBook((XSSFFlag) flag);

        Excel.Builder builder = new Excel.Builder(wb);
        builder.setSheetIndex(0);
        builder.setMetaDataRow(2);
        ExcelProcessor process = builder.build();

        System.out.println(Arrays.toString(process.getMetaData().toArray()));
    }
}
