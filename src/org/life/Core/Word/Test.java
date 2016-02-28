package org.life.Core.Word;


import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.life.Core.FileObjectFactory;
import org.life.Core.FileReaders;
import org.life.Core.TypeFlag.AbstractFlag.Flag;
import org.life.Core.TypeFlag.XWPFFlag;

import java.io.*;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Test {
    public static void main(String[] args) throws IOException {
        Flag flag = FileReaders.getFileFlag("E:\\101.docx");
        XWPFFlag f = (XWPFFlag)flag;
        XWPFDocument doc = FileObjectFactory.getDocument(f);

        List<XWPFTable> tables = doc.getTables();
        for(XWPFTable table: tables)
        {
            List<XWPFTableRow> rows = table.getRows();
            for(XWPFTableRow row: rows)
            {
                List<XWPFTableCell> cells = row.getTableCells();
                for(XWPFTableCell cell: cells)
                {
                    String data = cell.getText();
//                    data = data.replace("\\s", "");
                    data = data.replace("   ", "").replace(" ", "").replace(" ", "");
                    System.out.println(data);
                }
            }
        }

//        FileOutputStream file = new FileOutputStream("E:\\100.docx");
//        doc.write(file);
//        file.close();
//        f.close();
    }
}