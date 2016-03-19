package org.life.face;

import org.life.Core.FileReaders;
import org.life.Core.Word.Interface.Word;
import org.life.Core.Word.WordProcessor;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.*;
import java.util.List;
import java.util.function.Consumer;

public final class MyWindow extends JFrame {
    private java.util.List<Map<Integer, ? super JComponent>> headComponentGroup = new ArrayList<>();
    private Map<Integer, ? super JComponent> headSpecialComponentMap = new HashMap<>();
    private java.util.List<Map<Integer, ? super JComponent>> bodyComponentGroup = new ArrayList<>();
    private Map<Integer, ? super JComponent> footerComponentGroup = new HashMap<>();
    private Map<Integer, ? super JComponent> specialComponentMap = new HashMap<>();
    private java.util.List<JPanel> tabList = new ArrayList<>();
    private Map<String, Integer> directionMapper = new HashMap<>();

    private static final String RIGHT = "right";
    private static final String BOTTOM = "bottom";

    {
        directionMapper.put(RIGHT, WordProcessor.RIGHT);
        directionMapper.put(BOTTOM, WordProcessor.BOTTOM);
    }

    private JPanel head;
    private JPanel body;
    private JPanel footer;

    // 窗口大小
    private static final int WIDTH = 500;
    private static final int HEIGHT = 700;

    // head 控件组
    private static final int ADDRESS_FIELD = 0x00;
    private static final int IMPORT_BTN = 0x01;
    // Special
    private static final int SHEET_INDEX = 0x02;
    private static final int METADATA_ROW_FIELD = 0x03;
    private static final int LOAD_BUTTON = 0x04;
    private static final int READ_RADIO = 0x05;
    private static final int WRITE_RADIO = 0x06;
    private static final int SYMBOL_FILTER = 0x07;

    // tab 页控件组
    private static final int META_DATA_LABEL = 0x08;
    private static final int META_DATA_FIELD = 0x09;
    private static final int DATA_LENGTH_LABEL = 0x0A;
    private static final int DATA_LENGTH_FIELD = 0x0B;
    private static final int DIRECTION_COM_BOX = 0x0C;

    // footer 控件组
    private static final int OUTPUT_ADDRESS_FIELD = 0x0D;
    private static final int CHOOSE_ADDRESS_BTN = 0x0E;
    private static final int OUTPUT_BTN = 0x0F;

    // 特殊控件
    private static final int TAB_PANEL = 0x10;

    public MyWindow()
    {
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        setMinimumSize(new Dimension(WIDTH, HEIGHT));
        setLayout(new BorderLayout());
        setResizable(false);
        setTitle("Word & Excel 转换器");

        head = new JPanel();
        head.setLayout(new BoxLayout(head, BoxLayout.Y_AXIS));
        for(String unit: new String[]{"导入Excel", "导入Word"}) addHeadComponent(head, unit);
        addModelComponent(head);
        add(head, BorderLayout.NORTH);

        body = new JPanel();
        add(body, BorderLayout.CENTER);

        footer = new JPanel();
        addFooterComponent();
        add(footer, BorderLayout.SOUTH);

        // 监听器初始化
        new ListenRegister().init();
    }

    private void addModelComponent(JPanel head)
    {
        JLabel label;
        JTextField field;
        JButton loadButton = new JButton("加载模板");
        JRadioButton readRadio = new JRadioButton("只读模式");
        JRadioButton writeRadio = new JRadioButton("写模式");
        JPanel panelTmp = new JPanel();

        panelTmp.setLayout(new GridLayout(3, 4));

        label = new JLabel("Excel 表格索引号");
        panelTmp.add(label);

        field = new JTextField();
        panelTmp.add(field);
        headSpecialComponentMap.put(SHEET_INDEX, field);

        label = new JLabel("Excel 元数据行号");
        panelTmp.add(label);

        field = new JTextField();
        panelTmp.add(field);
        headSpecialComponentMap.put(METADATA_ROW_FIELD, field);

        label = new JLabel("Word 特殊字符过滤");
        panelTmp.add(label);

        field = new JTextField();
        panelTmp.add(field);
        headSpecialComponentMap.put(SYMBOL_FILTER, field);

        JLabel nbps = new JLabel("占位符");
        nbps.setVisible(false);
        panelTmp.add(nbps);

        nbps = new JLabel("占位符");
        nbps.setVisible(false);
        panelTmp.add(nbps);

        panelTmp.add(readRadio);
        headSpecialComponentMap.put(READ_RADIO, readRadio);

        panelTmp.add(writeRadio);
        headSpecialComponentMap.put(WRITE_RADIO, writeRadio);

        nbps = new JLabel("占位符");
        nbps.setVisible(false);
        panelTmp.add(nbps);

        panelTmp.add(loadButton);
        headSpecialComponentMap.put(LOAD_BUTTON, loadButton);

        head.add(panelTmp);
    }

    private void addHeadComponent(JPanel panel, String buttonName)
    {
        Map<Integer, ? super JComponent> map = new HashMap<>();
        JPanel panelTmp = new JPanel();
        JTextField textField;
        JButton button;

        textField = new JTextField();
        textField.setEditable(false);
        panelTmp.add(textField);
        map.put(ADDRESS_FIELD, textField);

        button = new JButton();
        button.setText(buttonName);
        map.put(IMPORT_BTN, button);

        panelTmp.setLayout(new BoxLayout(panelTmp, BoxLayout.X_AXIS));
        panelTmp.add(button);

        panel.add(panelTmp);
        headComponentGroup.add(map);
    }

    private void addTab(String... titles)
    {
        JTabbedPane tab;
        if(! specialComponentMap.containsKey(TAB_PANEL))
        {
            tab = new JTabbedPane();
            specialComponentMap.put(TAB_PANEL, tab);
        }
        else tab = ((JTabbedPane) specialComponentMap.get(TAB_PANEL));

        for(String unit: titles)
        {
            JPanel tabPanel = new JPanel();
            tabPanel.setLayout(new BoxLayout(tabPanel, BoxLayout.Y_AXIS));
            tab.add(new JScrollPane(tabPanel), unit);
            tabList.add(tabPanel);
        }

        body.setLayout(new BorderLayout());
        body.add(tab, BorderLayout.CENTER);
    }

    private void setComponentGroup(JPanel panel)
    {
        Map<Integer, ? super JComponent> group = new HashMap<>();
        JPanel panelTmp = new JPanel();
        JLabel label;
        JTextField textField;
        JComboBox<String> box;

        panelTmp.setLayout(new GridLayout());
        panelTmp.setMaximumSize(new Dimension(WIDTH, 30));

        label = new JLabel("字段");
        panelTmp.add(label);
        group.put(META_DATA_LABEL, new JLabel());

        textField = new JTextField();
        textField.setEditable(false);
        panelTmp.add(textField);
        group.put(META_DATA_FIELD, textField);

        label = new JLabel();
        label.setText("数据长度");
        panelTmp.add(label);
        group.put(DATA_LENGTH_LABEL, label);

        textField = new JTextField();
        panelTmp.add(textField);
        group.put(DATA_LENGTH_FIELD, textField);

        box = new JComboBox<>(new String[]{RIGHT, BOTTOM});
        panelTmp.add(box);
        group.put(DIRECTION_COM_BOX, box);

        panel.add(panelTmp);
        bodyComponentGroup.add(group);
    }

    private void addFooterComponent()
    {
        footer.setLayout(new BoxLayout(footer, BoxLayout.X_AXIS));

        JTextField textField = new JTextField();
        textField.setEditable(false);
        footer.add(textField);
        footerComponentGroup.put(OUTPUT_ADDRESS_FIELD, textField);

        JButton btn1 = new JButton("另存为");
        footer.add(btn1);
        footerComponentGroup.put(CHOOSE_ADDRESS_BTN, btn1);

        JButton btn2 = new JButton("生成Word");
        footer.add(btn2);
        footerComponentGroup.put(OUTPUT_BTN, btn2);
    }

    private final class ListenRegister {
        private void init()
        {
            headComponentListener();
            footerComponentListener();
        }

        private void headComponentListener()
        {
            // excel & word import action
            for(int x = 0; x < headComponentGroup.size(); x++)
            {
                final int index = x;
                ((JButton) headComponentGroup.get(index).get(IMPORT_BTN)).addActionListener(new ActionListener() {
                    @Override
                    public void actionPerformed(ActionEvent e) {
                        JFileChooser fileChooser = new JFileChooser();
                        fileChooser.setMultiSelectionEnabled(false);
                        fileChooser.showOpenDialog(new JPanel());
                        File file = fileChooser.getSelectedFile();

                        if(null == file)return;
                        ((JTextField) headComponentGroup.get(index).get(ADDRESS_FIELD)).setText(file.getAbsolutePath());
                    }
                });
            }

            // read & write model action
            ((JRadioButton) headSpecialComponentMap.get(READ_RADIO)).addActionListener(new ActionListener() {
                @Override
                public void actionPerformed(ActionEvent e) {
                    ((JRadioButton) headSpecialComponentMap.get(WRITE_RADIO)).setSelected(false);
                    for(Map<Integer, ? super JComponent> componentMap: bodyComponentGroup)
                        ((JTextField) componentMap.get(DATA_LENGTH_FIELD)).setEditable(true);
                }
            });
            ((JRadioButton) headSpecialComponentMap.get(WRITE_RADIO)).addActionListener(new ActionListener() {
                @Override
                public void actionPerformed(ActionEvent e) {
                    ((JRadioButton) headSpecialComponentMap.get(READ_RADIO)).setSelected(false);
                    for(Map<Integer, ? super JComponent> componentMap: bodyComponentGroup)
                        ((JTextField) componentMap.get(DATA_LENGTH_FIELD)).setEditable(false);
                }
            });
            ((JButton) headSpecialComponentMap.get(LOAD_BUTTON)).addActionListener(new ActionListener() {
                @Override
                public void actionPerformed(ActionEvent e) {
                    for(Map<Integer, ? super JComponent> componentMap: headComponentGroup)
                    {
                        String address = ((JTextField) componentMap.get(ADDRESS_FIELD)).getText();
                        if(address.equals(""))
                        {
                            JOptionPane.showMessageDialog(
                                    new JPanel(),
                                    "请选择模板",
                                    "Warning",
                                    JOptionPane.WARNING_MESSAGE);
                            return;
                        }
                    }

                    String sheetNum = ((JTextField) headSpecialComponentMap.get(SHEET_INDEX)).getText();
                    if(! sheetNum.matches("^\\d{1,2}$"))
                    {
                        JOptionPane.showMessageDialog(
                                new JPanel(),
                                "请输入正确的表索引号",
                                "Warning",
                                JOptionPane.WARNING_MESSAGE);
                        return;
                    }

                    String row = ((JTextField) headSpecialComponentMap.get(METADATA_ROW_FIELD)).getText();
                    if(! row.matches("^\\d{1,3}$"))
                    {
                        JOptionPane.showMessageDialog(
                                new JPanel(),
                                "请输入正确的行号",
                                "Warning",
                                JOptionPane.WARNING_MESSAGE);
                        return;
                    }

                    String excelPath = ((JTextField)headComponentGroup.get(0).get(ADDRESS_FIELD)).getText();
                    if(! FileReaders.checkSuffix(excelPath))
                    {
                        JOptionPane.showMessageDialog(
                                new JPanel(),
                                "请选择正确的Excel文件",
                                "Warning",
                                JOptionPane.WARNING_MESSAGE);
                        return;
                    }

                    // 清空所有所有的控件重新生成
                    if(Engines.isSameExcel(excelPath))return;
                    if(specialComponentMap.containsKey(TAB_PANEL))
                    {
                        ((JTabbedPane) specialComponentMap.get(TAB_PANEL)).removeAll();
                        tabList.clear();
                    }
                    if(bodyComponentGroup.size() > 0)bodyComponentGroup.clear();

                    Engines.initExcel(excelPath, Integer.parseInt(sheetNum) - 1, Integer.parseInt(row) - 1);
                    java.util.List<String> metaDataList = Engines.getExcelHandles().getMetaData();

                    if(null == metaDataList || metaDataList.size() <= 0)return;

                    // 利用 setVisible 重绘界面
                    body.setVisible(false);
                    addTab("table1");
                    bodyComponentListener(tabList.get(0), metaDataList);
                    body.setVisible(true);
                }
            });
        }

        private void bodyComponentListener(JPanel panel, java.util.List<String> metaDataList)
        {
            boolean isWrite = ((JRadioButton) headSpecialComponentMap.get(WRITE_RADIO)).isSelected();

            for(int x = 0; x < metaDataList.size(); x++) setComponentGroup(panel);
            for(int x = 0; x < bodyComponentGroup.size(); x++)
            {
                String text = metaDataList.get(x);
                ((JTextField) bodyComponentGroup.get(x).get(META_DATA_FIELD)).setText(text);

                if(isWrite)((JTextField) bodyComponentGroup.get(x).get(DATA_LENGTH_FIELD)).setEditable(false);
            }
        }

        private void footerComponentListener()
        {
            ((JButton) footerComponentGroup.get(CHOOSE_ADDRESS_BTN)).addActionListener(new ActionListener() {
                @Override
                public void actionPerformed(ActionEvent e) {
                    JFileChooser fileChooser = new JFileChooser();
                    fileChooser.setMultiSelectionEnabled(false);
                    fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                    fileChooser.showOpenDialog(new JPanel());
                    File file = fileChooser.getSelectedFile();

                    if(null == file)return;
                    ((JTextField) footerComponentGroup.get(OUTPUT_ADDRESS_FIELD)).setText(file.getAbsolutePath());
                }
            });

            ((JButton) footerComponentGroup.get(OUTPUT_BTN)).addActionListener(new ActionListener() {
                @Override
                public void actionPerformed(ActionEvent e) {
                    String excelPath = ((JTextField) headComponentGroup.get(0).get(ADDRESS_FIELD)).getText();
                    String wordPath = ((JTextField) headComponentGroup.get(1).get(ADDRESS_FIELD)).getText();
                    if(excelPath.equals("") || wordPath.equals("")) {
                        JOptionPane.showMessageDialog(
                                new JPanel(),
                                "请选择模板",
                                "Warning",
                                JOptionPane.WARNING_MESSAGE);
                        return;
                    }
                    if(! Engines.isSameExcel(excelPath)) {
                        JOptionPane.showMessageDialog(
                                new JPanel(),
                                "Excel 模板已改变, 请重新生成模板",
                                "Error",
                                JOptionPane.ERROR_MESSAGE);
                        return;
                    }

                    boolean isWrite = ((JRadioButton) headSpecialComponentMap.get(WRITE_RADIO)).isSelected();
                    boolean isRead = ((JRadioButton) headSpecialComponentMap.get(READ_RADIO)).isSelected();
                    if(! isRead && ! isWrite) {
                        JOptionPane.showMessageDialog(
                                new JPanel(),
                                "请选择只读或写模式",
                                "Warning",
                                JOptionPane.WARNING_MESSAGE);
                        return;
                    }

                    Map<String, Integer> metaDataAndLengthMap  = new HashMap<>();
                    Map<String, Integer> directionMap = new HashMap<>();

                    for(Map<Integer, ? super JComponent> bodyComponent: bodyComponentGroup) {
                        String metaData = ((JTextField) bodyComponent.get(META_DATA_FIELD)).getText();
                        Integer direction = directionMapper.get(((JComboBox) bodyComponent.get(DIRECTION_COM_BOX))
                                .getSelectedItem());

                        directionMap.put(metaData, direction);
                        if(isRead) {
                            String length = ((JTextField) bodyComponent.get(DATA_LENGTH_FIELD)).getText();

                            if(! length.matches("\\d{2}")) {
                                JOptionPane.showMessageDialog(
                                        new JPanel(),
                                        "无效的数据长度",
                                        "Warning",
                                        JOptionPane.WARNING_MESSAGE);
                                return;
                            }
                            else metaDataAndLengthMap.put(metaData, Integer.parseInt(length));
                        }
                    }

                    String[] symbol = ((JTextField) headSpecialComponentMap.get(SYMBOL_FILTER)).getText().split(",");
                    if(isWrite)Engines.initWordInWriteModel(wordPath, directionMap, symbol);
                    else Engines.initWordInReadModel(wordPath, directionMap, metaDataAndLengthMap, symbol);

                    String path = ((JTextField) footerComponentGroup.get(OUTPUT_ADDRESS_FIELD)).getText();
                    if(path.equals("")) {
                        JOptionPane.showMessageDialog(
                                new JPanel(),
                                "请选择保存路径",
                                "Warning",
                                JOptionPane.WARNING_MESSAGE);
                        return;
                    }

                    // Excel 转 Word 流程
                    List<Map<String, String>> data = Engines.getExcelHandles().getData();
                    Word word = Engines.getWordHandles();
                    int count = 1;
                    for(Map<String, String> unit: data) {
                        word.setTableData(1, unit);
                        word.saveAs(String.format("%s\\%d.docx", path, count++));
                    }

                    JOptionPane.showMessageDialog(new JPanel(), "转换完成");
                }
            });
        }
    }

    public static void main(String[] args)
    {
        MyWindow window = new MyWindow();
        window.setVisible(true);
    }
}
