package org.life.Core.Word.Interface;

import java.util.List;
import java.util.Map;

public interface WordProcessor {

    void writeText(String text);

    void setCellData(int tableNum, String fieldName, String data);

    void setTableData(Map<String, List<String>> map);

    String getCellData(int tableNum, String fieldName);

    Map<String, List<String>> getTableData(int tableNum);

    String readText();

    void close();
}