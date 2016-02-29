package org.life.Core.Word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.life.Core.Word.Interface.WordProcessor;

import java.util.*;

/**
 * Word 表格常用结构:
 * 一般分为 field -- value+ 与 field 两种基础结构
 *                              |
 *                            value+
 *
 * 策略：
 * 从A字段开始向A字段设定的方位开始循环读取空 TableCell, 直到遇到非空 XWPFTableCell 为止
 * 遇到一个 String 对应多个 XWPFTableCell 时, 对 String 转为 char[] 平均分到各个 TableCell 内
 *
 * 初始化时立刻对当前 Word 进行遍历, 获取所有的表及其字段, 使用 Map 组成一个 Field-XWPFTableCell 键值对结构
 * 同时对 TableIndex 进行关联, 把对 Word 的操作转换为对 Map 的操作
 * 一个 Word 对象仅处理一份 Word 文档, 由构建器保证
 *
 * dataDirectionMap 字段, 方位
 * tableMap 表索引号, field-TableCell Map(表结构)
 * cellMap 字段, XWPFTableCell对象
 */
public class Word implements WordProcessor {
    private Map<Integer, Map<String, List<XWPFTableCell>>> tableMap;
    private final Map<String, Integer> dataDirectionMap;
    private Map<String, Integer> metaDataAndLengthMap;
    private final List<String> trimSet;
    private final XWPFDocument doc;

    public static final int RIGHT = 0x00;
    public static final int BOTTOM = 0x01;

    private Word(Builder builder)
    {
        dataDirectionMap = builder.dataDirectionMap;
        metaDataAndLengthMap = builder.metaDataAndLengthMap;
        trimSet = builder.trimSet;
        doc = builder.doc;

        if(null == doc)throw new NullPointerException(String
                .format("%s: document Object is not init (constructor: Word).", getClass().getName()));
        if(null == dataDirectionMap || dataDirectionMap.size() == 0)throw new NullPointerException(String
                .format("%s: dataDirectionMap is not init (constructor: Word).", getClass().getName()));

        scanning$empty();
    }

    private void scanning$full()
    {
        int rowNum = 0;
        int cellNum = 0;
        int tableCount = 1;
        List<XWPFTable> tables = doc.getTables();

        for(XWPFTable table: tables)
        {
            Map<String, List<XWPFTableCell>> cellMap = new HashMap<>();
            tableMap.put(tableCount, cellMap);
            tableCount++;

            final int totalRow = table.getNumberOfRows();

            while(rowNum < totalRow)
            {
                XWPFTableCell cell = table.getRow(rowNum).getCell(cellNum);

                //  cell 为 null 代表当前 row 已空
                if(null == cell)
                {
                    rowNum++;
                    cellNum = 0;
                    continue;
                }

                if(! metaDataAndLengthMap.containsKey(cell.getText()))
                {
                    cellNum++;
                    continue;
                }

                // 到这里开始获取 key 并设定数据方向
                int direction = -1;
                String key = trim(cell.getText());

                // 此时的 rowNum 与 cellNum 均为 key 的方位！！！
                if(dataDirectionMap.containsKey(key)) direction = dataDirectionMap.get(key);

                scanning$core(table, key, cellMap, direction, rowNum, cellNum);
                cellNum++;
            }
        }
    }

    private void scanning$core()
    {

    }

    private void scanning$empty()
    {
        int rowNum = 0;
        int cellNum = 0;
        int tableCount = 1;
        List<XWPFTable> tables = doc.getTables();

        for(XWPFTable table: tables)
        {
            Map<String, List<XWPFTableCell>> cellMap = new HashMap<>();
            tableMap.put(tableCount, cellMap);
            tableCount++;

            final int totalRow = table.getNumberOfRows();

            while(rowNum < totalRow)
            {
                XWPFTableCell cell = table.getRow(rowNum).getCell(cellNum);

                //  cell 为 null 代表当前 row 已空
                if(null == cell)
                {
                    rowNum++;
                    cellNum = 0;
                    continue;
                }

                // cell 为空时跳过本次训话
                if(cell.getText().equals(""))
                {
                    cellNum++;
                    continue;
                }

                // 到这里开始获取 key 并设定数据方向
                int direction = -1;
                String key = trim(cell.getText());

                // 此时的 rowNum 与 cellNum 均为 key 的方位！！！
                if(dataDirectionMap.containsKey(key)) direction = dataDirectionMap.get(key);

                scanning$core(table, key, cellMap, direction, rowNum, cellNum);
                cellNum++;
            }
        }
    }

    private void scanning$core(XWPFTable table, String key, Map<String, List<XWPFTableCell>> cellMap,
                               int direction, int rowNumCopy, int cellNumCopy)
    {
        while(true)
        {
            if(BOTTOM == direction)rowNumCopy++;
            else if(RIGHT == direction)cellNumCopy++;
            else return;

            XWPFTableCell cell;
            try {
                cell = table.getRow(rowNumCopy).getCell(cellNumCopy);
                if(null == cell || (! cell.getText().equals("")))return;
            }
            catch (NullPointerException e) {
                return;
            }

            if(! cellMap.containsKey(key))
            {
                List<XWPFTableCell> tmp = new ArrayList<>();
                tmp.add(cell);
                cellMap.put(key, tmp);
            }
            else cellMap.get(key).add(cell);
        }
    }

    static class Builder {
        private Map<String, Integer> dataDirectionMap;
        private Map<String, Integer> metaDataAndLengthMap;
        private List<String> trimSet;
        private XWPFDocument doc;

        /**
         * 初始化为 2007或以上 Word文档类
         * @param xwpfDocument Word操作对象
         */
        public Builder(XWPFDocument xwpfDocument)
        {
            doc = xwpfDocument;
        }

        /**
         * 设置字段数据方向, 可选的方向有 Word.RIGHT 与 Word.BOTTOM
         * @param dataDirectionMap 一个字段数据方向的集合
         */
        public void setDataDirectionMap(Map<String, Integer> dataDirectionMap) {
            this.dataDirectionMap = dataDirectionMap;
        }

        /**
         * 设置元数据及其数据量, 设定后表格扫描器将改变扫描风格, 扫描所有单元格直到满足指定数据量
         * @param metaDataAndLengthMap 一个包含元数据及其数据量的 Map, 其中 key 为元数据, value 为数据量
         */
        public void setMetaDataAndLengthMap(Map<String, Integer> metaDataAndLengthMap) {
            this.metaDataAndLengthMap = metaDataAndLengthMap;
        }

        /**
         * 声明文档中所有不同类型的空格, 制表符或未知但表现为空格的分隔符, 设置后将会过滤所有声明的分隔符
         * @param symbols 一个或多个表现为空格的分隔符
         */
        public void setTrimSet(String... symbols)
        {
            trimSet = Arrays.asList(symbols);
        }

        public Word Build()
        {
            return new Word(this);
        }
    }

    @Override
    public void writeText(String text) {

    }

    @Override
    public void setCellData(int tableNum, String fieldName, String data)
    {
        List<XWPFTableCell> cellList;

        try {
            cellList = tableMap.get(tableNum).get(fieldName);
        }
        catch (NullPointerException e) {
            return;
        }

        if(cellList.size() <= 0)throw new IndexOutOfBoundsException(String
                .format("%s: size is lower of zero: %d (Method: setCellData)", getClass().getName(), cellList.size()));

        if(cellList.size() == 1) cellList.get(0).setText(data);
        else
        {
            char[] arr = data.toCharArray();
            int average = arr.length / cellList.size();

            for(int x = 0; x < cellList.size(); x++)
                cellList.get(x).setText(new String(arr, x * average, average));
        }
    }

    @Override
    public void setTableData(int tableNum, Map<String, String> map)
    {
        Set<Map.Entry<String, String>> entrySet = map.entrySet();
        for(Map.Entry<String, String> entry: entrySet) setCellData(tableNum, entry.getKey(), entry.getValue());
    }

    @Override
    public String getCellData(int tableNum, String fieldName)
    {
        return null;
    }

    @Override
    public Map<String, List<String>> getTableData(int tableNum) {
        return null;
    }

    @Override
    public String readText() {
        return null;
    }

    @Override
    public void close() {

    }

    private String trim(String tar)
    {
        for(String unit: trimSet)tar = tar.replace(unit, "");
        return tar;
    }
}
