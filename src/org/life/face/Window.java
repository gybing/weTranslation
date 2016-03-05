package org.life.face;

import javax.swing.*;
import java.awt.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Window extends JFrame{
    private Map<String, ? super JComponent> componentMap = new HashMap<>();

    private static String OPEN_EXCEL_BTN = "openExcelBtn";

    public Window()
    {
        setSize(800, 600);
        setLayout(new BorderLayout());
        componentMap.put(OPEN_EXCEL_BTN, new JButton("导入Excel"));

        add((JButton) componentMap.get(OPEN_EXCEL_BTN), BorderLayout.NORTH);
    }

    public static void main(String[] args)
    {
        Window window = new Window();
        window.setVisible(true);
    }
}
