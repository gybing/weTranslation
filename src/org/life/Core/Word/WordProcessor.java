package org.life.Core.Word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.life.Core.Word.Interface.Word;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;

/**
 * Word 表格常用结构:
 * 一般分为 field -- value+ 与 field 两种基础结构
 *                              |
 *                            value+
 *
 * 策略：
 * 从A字段开始向A字段设定的方位开始循环读取空 XWPFTableCell, 直到遇到非空 XWPFTableCell 为止
 * 遇到一个 String 对应多个 XWPFTableCell 时, 对 String 转为 char[] 平均分到各个 XWPFTableCell 内
 *
 * 初始化时立刻对当前 Word 进行遍历, 获取所有的表及其字段, 使用 Map 组成一个 Field-XWPFTableCell 键值对结构
 * 同时对 TableIndex 进行关联, 把对 Word 的操作转换为对 Map 的操作
 * 一个 Word 对象仅处理一份 Word 文档, 由构建器保证
 *
 * tableMap 表格Map，一个 entry 便是一个表格
 * dataDirectionMap 字段方位，key-字段 value-方位 ，决定字段读取方向
 * metaDataAndLengthMap 元数据及其数据长度， key-元数据 value-数据长度，非 null 时为读取模式
 * trimSet 过滤字符集合 将会过滤掉在此列表中所有的字符
 */
public class WordProcessor implements Word {
    private Map<Integer, Map<String, List<XWPFTableCell>>> tableMap;
    private final Map<String, Integer> dataDirectionMap;
    private Map<String, Integer> metaDataAndLengthMap;
    private final List<String> trimSet;
    private final XWPFDocument doc;

    public static final int RIGHT = 0x00;
    public static final int BOTTOM = 0x01;

    private WordProcessor(Builder builder)
    {
        dataDirectionMap = builder.dataDirectionMap;
        metaDataAndLengthMap = builder.metaDataAndLengthMap;
        trimSet = builder.trimSet;
        doc = builder.doc;
        tableMap = new HashMap<>();

        if(null == doc)throw new NullPointerException(String
                .format("%s: document Object is not init (constructor: Word).", getClass().getName()));
        if(null == dataDirectionMap || dataDirectionMap.size() == 0)throw new NullPointerException(String
                .format("%s: dataDirectionMap is not init (constructor: Word).", getClass().getName()));

        scanning();
    }

    private void scanning()
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

                // 到这里开始获取 key 并设定数据方向
                int direction = -1;
                String key = trim(cell.getText());

                // 此时的 rowNum 与 cellNum 均为 key 的方位！！！
                if(dataDirectionMap.containsKey(key)) direction = dataDirectionMap.get(key);
                if(direction != -1)scanning$core(table, key, cellMap, direction, rowNum, cellNum);

                cellNum++;
            }
        }
    }

    private void scanning$core(XWPFTable table, String key, Map<String, List<XWPFTableCell>> cellMap,
                               int direction, int rowNumCopy, int cellNumCopy)
    {
        int length = -1;
        if(metaDataAndLengthMap != null)length = metaDataAndLengthMap.get(key);

        while(true)
        {
            if(BOTTOM == direction)rowNumCopy++;
            else if(RIGHT == direction)cellNumCopy++;
            else return;

            XWPFTableCell cell;
            try {
                cell = table.getRow(rowNumCopy).getCell(cellNumCopy);
                if(null == cell)return;
                if(metaDataAndLengthMap != null && length <= 0)return;
                if(null == metaDataAndLengthMap && (! cell.getText().equals("")))return;
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

            if(length != -1)length--;
        }
    }

    public static final class Builder {
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
        public Builder setDataDirectionMap(Map<String, Integer> dataDirectionMap)
        {
            this.dataDirectionMap = dataDirectionMap;
            return this;
        }

        /**
         * 设置元数据及其数据量, 设定后表格扫描器将改变扫描风格, 扫描所有单元格直到满足指定数据量
         * @param metaDataAndLengthMap 一个包含元数据及其数据量的 Map, 其中 key 为元数据, value 为数据量
         */
        public Builder setMetaDataAndLengthMap(Map<String, Integer> metaDataAndLengthMap)
        {
            this.metaDataAndLengthMap = metaDataAndLengthMap;
            return this;
        }

        /**
         * 声明文档中所有不同类型的空格, 制表符或未知但表现为空格的分隔符, 设置后将会过滤所有声明的分隔符
         * @param symbols 一个或多个表现为空格的分隔符
         */
        public Builder setTrimSet(String... symbols)
        {
            trimSet = Arrays.asList(symbols);
            return this;
        }

        public WordProcessor build()
        {
            return new WordProcessor(this);
        }
    }

    /**
     * 暂未实现
     * @param text 需要写入文档的文字
     */
    @Override
    public void writeText(String text) {

    }

    /**
     * 更改 Word 中指定表格的字段的数据
     * @param tableNum 表索引号(按 Word 中表格的顺序)
     * @param fieldName 字段名
     * @param data 新数据
     */
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

    /**
     * 批量更改 Word 中指定表名的字段的数据
     * @param tableNum 表索引号(按 Word 中表格的顺序)
     * @param map 新数据集合
     */
    @Override
    public void setTableData(int tableNum, Map<String, String> map)
    {
        Set<Map.Entry<String, String>> entrySet = map.entrySet();
        for(Map.Entry<String, String> entry: entrySet) setCellData(tableNum, entry.getKey(), entry.getValue());
    }

    @Override
    public String getCellData(int tableNum, String fieldName)
    {
        if(tableNum >= tableMap.size())throw new IndexOutOfBoundsException(String
                .format("%s: out of tableIndex: %d (Method: getCellData)", getClass().getName(), tableNum));

        Map<String, List<XWPFTableCell>> table = tableMap.get(tableNum);
        List<XWPFTableCell> dlist = table.get(fieldName);

        if(null == dlist)return "";

        StringBuilder strBuilder = new StringBuilder();
        for(XWPFTableCell cell: dlist)strBuilder.append(cell.getText());
        return strBuilder.toString();
    }

    /**
     * 获取 Word 中指定表的所有数据
     * @param tableNum 表索引号(按 Word 中表格的顺序)
     * @return 返回一个包含指定表的数据的集合
     */
    @Override
    public Map<String, String> getTableData(int tableNum)
    {
        if(tableNum >= tableMap.size())throw new IndexOutOfBoundsException(String
                .format("%s: out of tableIndex: %d (Method: getCellData)", getClass().getName(), tableNum));

        Map<String, String> map = new HashMap<>();
        Map<String, List<XWPFTableCell>> mapTmp = tableMap.get(tableNum);

        Set<Map.Entry<String, List<XWPFTableCell>>> entrySet = mapTmp.entrySet();
        for(Map.Entry<String, List<XWPFTableCell>> unit: entrySet)
        {
            List<XWPFTableCell> dList = unit.getValue();
            StringBuilder strBuilder = new StringBuilder();
            for(XWPFTableCell cell: dList)strBuilder.append(cell.getText());
            map.put(unit.getKey(), strBuilder.toString());
        }

        return map;
    }

    /**
     * 获取 Word 文章(空实现)
     * @return 返回 Word 所有文章段落
     */
    @Override
    public String readText() {
        return null;
    }

    /**
     * 当前 Word 另存为
     * @param path 新文件路径
     */
    @Override
    public void saveAs(String path)
    {
        try {
            saveAs(new FileOutputStream(path));
        }
        catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 当前 Excel 另存为
     * @param outputStream 新文件路径
     */
    @Override
    public void saveAs(OutputStream outputStream)
    {
        try {
            doc.write(outputStream);
        }
        catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 关闭资源
     */
    @Override
    public void close()
    {
        try {
            doc.close();
        }
        catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private String trim(String tar)
    {
        for(String unit: trimSet)tar = tar.replace(unit, "");
        return tar;
    }
}
