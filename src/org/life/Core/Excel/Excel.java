package org.life.Core.Excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.life.Core.Excel.Interface.ExcelProcessor;
import org.life.Exception.ColumnOutOfBoundsException;
import org.life.Exception.KeyException;
import org.life.Exception.RowOutOfBoundsException;

import java.io.IOException;
import java.util.*;


public class Excel implements ExcelProcessor{
    private Map<String, Integer> mappingMap;
    private List<String> metaDataList;
    private final String sheetName;
    private final int sheetIndex;
    private final int metaDataRow;
    private final Workbook wb;

    {
        mappingMap = new HashMap<>();
        metaDataList = new ArrayList<>();
    }

    private Excel(Builder builder)
    {
        sheetName = builder.sheetName;
        sheetIndex = builder.sheetIndex;
        metaDataRow = builder.metaDataRow;
        wb = builder.wb;
        startInit();
    }

    private void startInit()
    {
        Sheet sheet;
        if(null != sheetName)sheet = wb.getSheet(sheetName);
        else sheet = wb.getSheetAt(sheetIndex);

        Row row = sheet.getRow(metaDataRow);
        for(int x = 0, offset = 0; x < row.getLastCellNum(); x++)
        {
            Cell cell = row.getCell(x);
            String data = cell.getStringCellValue();
            if(data.equals(""))
            {
                offset++;
                continue;
            }

            mappingMap.put(data, x + offset);
            metaDataList.add(data);
        }
    }

    /**
     * SheetName 或 SheetIndex 必须指定一个
     * metaDataRow 为元数据行, 必须指定该字段
     */
    static class Builder {
        private String sheetName;
        private int sheetIndex;
        private int metaDataRow;
        private Workbook wb;

        /**
         * 初始化为 97-2003 Excel文档类
         * @param hssfWorkbook Excel操作对象
         */
        public Builder(HSSFWorkbook hssfWorkbook)
        {
            wb = hssfWorkbook;
        }

        /**
         * 初始化为 2007或以上 Excel文档类
         * @param xssfWorkbook Excel操作对象
         */
        public Builder(XSSFWorkbook xssfWorkbook)
        {
            wb = xssfWorkbook;
        }

        /**
         * 设定需要处理的表名
         * @param sheetName 表名
         */
        public void setSheetName(String sheetName) {
            this.sheetName = sheetName;
        }

        /**
         * 设定需要处理的表的索引位置
         * @param sheetIndex 索引号
         */
        public void setSheetIndex(int sheetIndex) {
            this.sheetIndex = sheetIndex;
        }

        /**
         * 设定元数据的起始行号
         * @param metaDataRow 行号
         */
        public void setMetaDataRow(int metaDataRow) {
            this.metaDataRow = metaDataRow;
        }

        public ExcelProcessor build()
        {
            return new Excel(this);
        }
    }

    /**
     * 更改指定单元格的数据
     * @param newData 新数据
     * @param colName 列名
     * @param rowNumber 行数, 如果 rowNumber <= 0 或 rowNumber > lastRowNum 将引发 RowOutOfBoundsException
     */
    @Override
    public void setCellData(String newData, String colName, int rowNumber)
    {
        if(! mappingMap.containsKey(colName))throw new KeyException(String
                .format("%s: ColumnNotExist: %s (Method: setCellData).", getClass().getName(), colName));

        setCellData(newData, mappingMap.get(colName), rowNumber);
    }

    /**
     * 更改指定的单元格数据
     * @param newData 新数据
     * @param colNumber 列数, 如果 colNumber < 0 将引发 ColumnOutOfBoundsException
     * @param rowNumber 行数, 如果 rowNumber <= 0 或 rowNumber >= lastRowNum 将引发 RowOutOfBoundsException
     */
    @Override
    public void setCellData(String newData, int colNumber, int rowNumber)
    {
        if(rowNumber <= metaDataRow)throw new RowOutOfBoundsException(String
                .format("%s: OutOfIndex: %d (Method: setCellData).", getClass().getName(), rowNumber));
        if(colNumber < 0)throw new ColumnOutOfBoundsException(String
                .format("%s: OutOfIndex: %d (Method: setCellData).", getClass().getName(), colNumber));

        Sheet sheet = wb.getSheet(sheetName);
        if(rowNumber >= sheet.getLastRowNum())throw new RowOutOfBoundsException(String
                .format("%s: OutOfIndex: %d (Method: setCellData).", getClass().getName(), rowNumber));

        Row row = sheet.getRow(rowNumber);
        row.getCell(colNumber).setCellValue(newData);
    }

    /**
     * 在指定行内修改数据, 仅修改数据集合中存在的列名所对应的数据
     * @param data 数据集合, 其中 key 为修改的列名, value 为新数据
     * @param rowNumber 行数
     */
    @Override
    public void setRow(Map<String, String> data, int rowNumber)
    {
        if(rowNumber <= metaDataRow)throw new RowOutOfBoundsException(String
                .format("%s: OutOfIndex: %d (Method: setRow).", getClass().getName(), rowNumber));
        if(null == data)throw new NullPointerException(String
                .format("%s: Object is Null: data (Method: setRow)", getClass().getName()));

        // 元数据匹配检查
        String result = metaDataChecker(data.keySet());
        if(! result.equals(""))throw new KeyException(String
                .format("%s: ColumnNotExist: %s (Method: setRow)", getClass().getName(), result));

        Sheet sheet = wb.getSheet(sheetName);
        if(rowNumber >= sheet.getLastRowNum())throw new RowOutOfBoundsException(String
                .format("%s: OutOfIndex: %d (Method: setRow).", getClass().getName(), rowNumber));

        Row row = sheet.getRow(rowNumber);
        Set<Map.Entry<String, String>> iter = data.entrySet();

        for(Map.Entry<String, String> entry: iter)
            row.getCell(mappingMap.get(entry.getKey())).setCellValue(entry.getValue());
    }

    /**
     * 增加一行新的数据, 仅增加数据集合中存在的列名所对应的数据
     * @param data 数据集合, 其中 key 为修改的列名, value 为新数据
     */
    @Override
    public void addRow(Map<String, String> data)
    {
        if(null == data)throw new NullPointerException(String
                .format("%s: Object is Null: data (Method: addRow)", getClass().getName()));

        String result = metaDataChecker(data.keySet());
        if(! result.equals(""))throw new KeyException(String
                .format("%s: ColumnNotExist: %s (Method: addRow)", getClass().getName(), result));

        Sheet sheet = wb.getSheet(sheetName);
        Row row = sheet.createRow(sheet.getLastRowNum());
        Set<Map.Entry<String, String>> iter = data.entrySet();

        for(Map.Entry<String, String> entry: iter)
            row.getCell(mappingMap.get(entry.getKey())).setCellValue(entry.getValue());
    }

    /**
     * 获取元数据
     * @return 返回一个元数据列表
     */
    @Override
    public List<String> getMetaData() {
        return metaDataList;
    }

    /**
     * 获取当前表格的所有数据, 其中 key 为元数据, value 为数据, 每行数据为一个 Map 集合
     * @return 返回一个数据列表
     */
    @Override
    public List<Map<String, String>> getData()
    {
        Sheet sheet = wb.getSheet(sheetName);
        List<Map<String, String>> dataList = new ArrayList<>();

        for(int x = 1; x < sheet.getLastRowNum(); x++)
        {
            Row row = sheet.getRow(x);
            Map<String, String> dataMap = new HashMap<>();
            Set<Map.Entry<String, Integer>> iter = mappingMap.entrySet();

            for(Map.Entry<String, Integer> entry: iter)
            {
                String data = row.getCell(entry.getValue()).getStringCellValue();
                dataMap.put(entry.getKey(), data);
            }

            dataList.add(dataMap);
        }

        return dataList;
    }

    /**
     * 获取指定行的所有数据, 其中 key 为元数据, value 为数据
     * @param rowNumber 行数
     * @return 返回一个数据列表
     */
    @Override
    public Map<String, String> getRowData(int rowNumber)
    {
        if(rowNumber <= metaDataRow)throw new RowOutOfBoundsException(String
                .format("%s: OutOfIndex: %d (Method: getRowData).", getClass().getName(), rowNumber));

        Sheet sheet = wb.getSheet(sheetName);
        Row row = sheet.getRow(rowNumber);
        Map<String, String> dataMap = new HashMap<>();

        if(rowNumber >= sheet.getLastRowNum())throw new RowOutOfBoundsException(String
                .format("%s: OutOfIndex: %d (Method: getRowData).", getClass().getName(), rowNumber));

        Set<Map.Entry<String, Integer>> iter = mappingMap.entrySet();
        for(Map.Entry<String, Integer> entry: iter)
        {
            String data = row.getCell(entry.getValue()).getStringCellValue();
            dataMap.put(entry.getKey(), data);
        }

        return dataMap;
    }

    /**
     * 获取指定单元格数据
     * @param colName 列名
     * @param rowNumber 行数
     * @return 返回一个单元格数据
     */
    @Override
    public String getCellData(String colName, int rowNumber)
    {
        if(null == colName || (! mappingMap.containsKey(colName)))throw new KeyException(String
                .format("%s: ColumnNotExist: %s (Method: getCellData).", getClass().getName(), colName));

        return getCellData(mappingMap.get(colName), rowNumber);
    }

    /**
     * 获取指定单元格数据
     * @param colNumber 列数
     * @param rowNumber 行数
     * @return 返回一个单元格数据
     */
    @Override
    public String getCellData(int colNumber, int rowNumber)
    {
        if(rowNumber <= metaDataRow)throw new RowOutOfBoundsException(String
                .format("%s: OutOfIndex: %d (Method: getCellData).", getClass().getName(), rowNumber));
        if(colNumber < 0)throw new ColumnOutOfBoundsException(String
                .format("%s: OutOfIndex: %d (Method: getCellData).", getClass().getName(), colNumber));

        Sheet sheet = wb.getSheet(sheetName);
        if(rowNumber >= sheet.getLastRowNum())throw new RowOutOfBoundsException(String
                .format("%s: OutOfIndex: %d (Method: getCellData).", getClass().getName(), rowNumber));

        Row row = sheet.getRow(rowNumber);
        return row.getCell(colNumber).getStringCellValue();
    }

    /**
     * 关闭资源
     */
    @Override
    public void close()
    {
        try
        {
            wb.close();
        }
        catch (IOException e)
        {
            e.printStackTrace();
            throw new RuntimeException(String
                    .format("%s: close failure. (Method: close)", getClass().getName()));
        }
    }

    /**
     * key / metaData 匹配检查
     * @param keySet 数据的key的set集合
     * @return 若所有key都与元数据适配, 返回空字符串 ,否则返回不匹配的 key
     */
    private String metaDataChecker(Set<String> keySet)
    {
        for(String unit: keySet)
        {
            if(! metaDataList.contains(unit))return unit;
        }
        return "";
    }
}