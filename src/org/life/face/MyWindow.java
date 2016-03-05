package org.life.face;

import org.life.Core.Word.Word;

import javax.swing.*;
import java.awt.*;
import java.util.*;

public class MyWindow extends JFrame{
    private java.util.List<Map<Integer, ? super JComponent>> headComponentGroup = new ArrayList<>();
    private java.util.List<Map<Integer, ? super JComponent>> bodyComponentGroup = new ArrayList<>();
    private Map<Integer, ? super JComponent> footerComponentGroup = new HashMap<>();
    private java.util.List<JPanel> tabList = new ArrayList<>();

    private JPanel head;
    private JPanel body;
    private JPanel footer;

    // 窗口大小
    private static final int WIDTH = 400;
    private static final int HEIGHT = 500;

    // head 控件组
    private static final int ADDRESS_FIELD = 0x00;
    private static final int IMPOER_BTN = 0x01;

    // tab 页控件组
    private static final int META_DATA_LABEL = 0x02;
    private static final int META_DATA_FIELD = 0x03;
    private static final int DATA_LENGTH_LABEL = 0x04;
    private static final int DATA_LENGTH_FIELD = 0x05;
    private static final int DIRECTION_COM_BOX = 0x06;

    // footer 控件组
    private static final int OUTPUT_ADDRESS_FIELD = 0x07;
    private static final int CHOOSE_ADDRESS_BTN = 0x08;
    private static final int OUTPUT_BTN = 0x09;

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
        add(head, BorderLayout.NORTH);

        body = new JPanel();
        addTab("table1", "table2", "table3");
        for(int x = 0; x < 10; x++) getComponentGroup(tabList.get(0));
        add(body, BorderLayout.CENTER);

        footer = new JPanel();
        addFooterComponent();
        add(footer, BorderLayout.SOUTH);
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
        map.put(IMPOER_BTN, button);

        panelTmp.setLayout(new BoxLayout(panelTmp, BoxLayout.X_AXIS));
        panelTmp.add(button);

        panel.add(panelTmp);
        headComponentGroup.add(map);
    }

    private void addTab(String... titles)
    {
        JTabbedPane tab = new JTabbedPane();

        for(String unit: titles)
        {
            JPanel tabPanel = new JPanel();
            tabPanel.setLayout(new BoxLayout(tabPanel, BoxLayout.Y_AXIS));
            tab.add(tabPanel, unit);
            tabList.add(tabPanel);
        }

        body.setLayout(new BorderLayout());
        body.add(tab, BorderLayout.CENTER);
    }

    private void getComponentGroup(JPanel panel)
    {
        Map<Integer, ? super JComponent> group = new HashMap<>();
        JPanel panelTmp = new JPanel();
        JLabel label;
        JTextField textField;
        JComboBox<Integer> box;

        panelTmp.setLayout(new BoxLayout(panelTmp, BoxLayout.X_AXIS));
        panelTmp.setMaximumSize(new Dimension(WIDTH, 30));

        label = new JLabel();
        label.setText("字段");
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

        box = new JComboBox<>(new Integer[]{Word.RIGHT, Word.BOTTOM});
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

    public static void main(String[] args)
    {
        MyWindow window = new MyWindow();
        window.setVisible(true);
    }
}
