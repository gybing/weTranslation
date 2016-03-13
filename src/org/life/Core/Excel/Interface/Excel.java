package org.life.Core.Excel.Interface;

import java.io.OutputStream;
import java.util.List;
import java.util.Map;

public interface Excel {

    void setCellData(String newData, String colName, int rowNumber);

    void setCellData(String newData, int colNumber, int rowNumber);

    void setRow(Map<String, String> data, int rowNumber);

    void addRow(Map<String, String> data);

    List<String> getMetaData();

    List<Map<String, String>> getData();

    Map<String, String> getRowData(int rowNumber);

    String getCellData(String colName, int rowNumber);

    String getCellData(int colNumber, int rowNumber);

    void saveAs(String path);

    void saveAs(OutputStream outputStream);

    void close();
}
