package org.life.face;

import javax.swing.*;

public class weFace extends JFrame{
    private JTextField excelHref;
    private JButton btnWord;
    private JButton btnExcel;
    private JTabbedPane tablePane;
    private JTextField textField2;
    private JButton 生成WordButton;
    private JTextField mdField1;
    private JComboBox directionBox1;
    private JTextField dLengthField1;
    private JTextField mdField2;
    private JTextField dLengthField2;
    private JComboBox directionBox2;
    private JTextField textField7;
    private JTextField textField8;
    private JComboBox comboBox3;
    private JComboBox comboBox4;
    private JTextField textField9;
    private JTextField textField10;
    private JRadioButton 填充模式RadioButton;
    private JRadioButton 读取模式RadioButton;
    private JButton button1;
    private JTextField wordHref;
    private JLabel mdLabel1;
    private JLabel dlLabel1;
    private JLabel mdLabel2;
    private JLabel dlLabel2;
    private JPanel panel2;
    private JPanel panel1;
    private JPanel table1;
    private JPanel table2;
    private JPanel table3;
    private JPanel table4;
    private JLabel label;
    private JLabel abc;
    private JLabel wdi;
    private JLabel wdaw;

    private void createUIComponents() {
        // TODO: place custom component creation code here
    }

    public static void main(String[] args)
    {
        JFrame frame = new JFrame();
        frame.setContentPane(new weFace().panel1);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);
    }
}
