package org.life.Core.Word.Interface;

import java.io.OutputStream;
import java.util.List;
import java.util.Map;

public interface WordProcessor {

    void writeText(String text);

    void setCellData(int tableNum, String fieldName, String data);

    void setTableData(int tableNum, Map<String, String> map);

    String getCellData(int tableNum, String fieldName);

    Map<String, String> getTableData(int tableNum);

    String readText();

    void saveAs(String path);

    void saveAs(OutputStream outputStream);

    void close();
}