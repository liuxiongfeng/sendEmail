
import javafx.embed.swing.JFXPanel;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

import javax.swing.*;
import javax.swing.filechooser.FileFilter;

public class SendMail extends JFrame {
    String absolutePath = "";
    String filepath1 = "";
    String filepath2 = "";
    String picpath = "";

    public static void main(String args[]) {
        try {
            JFrame.setDefaultLookAndFeelDecorated(true);
            UIManager.setLookAndFeel("ch.randelshofer.quaqua.QuaquaLookAndFeel");
            new SendMail();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }


    public SendMail() {
        final JFrame jf=this;
        final JPanel jp = new JPanel(null);
        jp.setBounds(0,0,800,700);
        Container contentPane = getContentPane();
        JLabel labelIP = new JLabel("主题:");
        labelIP.setBounds(300, 100, 100, 30);

        final JTextField jtfIP = new JTextField("邮件标题");
        jtfIP.setBounds(400, 100, 200, 30);

        JLabel labelDatabaseName = new JLabel("正文:");
        labelDatabaseName.setBounds(300, 200, 100, 30);

        final JTextField jtfDatabaseName = new JTextField("正文内容");
        jtfDatabaseName.setBounds(400, 200, 200, 30);

        JLabel labelUserName = new JLabel("附件1:");
        labelUserName.setBounds(300, 300, 100, 30);
        final JButton fujian1 = new JButton("选择附件");
        fujian1.setBounds(400, 300, 200, 30);
        /*final JTextField jtfUserName = new JTextField("root");
        jtfUserName.setBounds(400, 300, 200, 30);*/

        JLabel labelPassword = new JLabel("附件2:");
        labelPassword.setBounds(300, 350, 100, 30);
        final JButton fujian2 = new JButton("选择附件");
        fujian2.setBounds(400, 350, 200, 30);
        /*final JTextField jtfPassword = new JTextField("root");
        jtfPassword.setBounds(400, 400, 200, 30);*/

        JLabel labelPic = new JLabel("图片:");
        labelPassword.setBounds(300, 500, 100, 30);
        final JButton tupian = new JButton("选择图片");
        fujian2.setBounds(400, 500, 100, 30);

        final JLabel importing = new JLabel("导入中...,请稍后");
        importing.setFont(new   java.awt.Font("Dialog",   1,   25));
        importing.setForeground(Color.RED);
        importing.setBounds(400, 230, 200, 70);
        importing.setVisible(false);

        JButton open = new JButton("邮箱文件");
        open.setBounds(300, 600, 100, 30);

        JButton submit = new JButton("确定");
        submit.setBounds(500, 600, 100, 30);

        this.setLayout(null);

        //设置尺寸
        jp.add(labelIP);
        jp.add(jtfIP);
        jp.add(labelDatabaseName);
        jp.add(jtfDatabaseName);
        jp.add(labelUserName);
        jp.add(fujian1);
        jp.add(labelPassword);
        jp.add(fujian2);
        jp.add(labelPic);
        jp.add(tupian);
        jp.add(importing);
        jp.add(open);
        jp.add(submit);
        contentPane.add(jp);
        this.setBounds(0, 0, 800, 700);
        this.setVisible(true);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        fujian1.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser jfc = new JFileChooser();
                jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
                jfc.setFileFilter(new FileFilter() {
                    @Override
                    public boolean accept(File file) {
                        String name = file.getName();
                        return file.isDirectory() || name.toLowerCase().endsWith(".xls") || name.toLowerCase().endsWith(".xlsx") || name.toLowerCase().endsWith(".word") || name.toLowerCase().endsWith(".pdf");  // 仅显示目录和xls、xlsx文件
                    }

                    @Override
                    public String getDescription() {
                        return "*.xls;*.xlsx;*.word;*.pdf";
                    }
                });
                jfc.showDialog(new JLabel(), "选择附件1");
                filepath1 = jfc.getSelectedFile().getAbsolutePath();
            }
        });
        fujian2.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser jfc = new JFileChooser();
                jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
                jfc.setFileFilter(new FileFilter() {
                    @Override
                    public boolean accept(File file) {
                        String name = file.getName();
                        return file.isDirectory() || name.toLowerCase().endsWith(".xls") || name.toLowerCase().endsWith(".xlsx") || name.toLowerCase().endsWith(".word") || name.toLowerCase().endsWith(".pdf");  // 仅显示目录和xls、xlsx文件
                    }

                    @Override
                    public String getDescription() {
                        return "*.xls;*.xlsx;*.word;*.pdf";
                    }
                });
                jfc.showDialog(new JLabel(), "选择附件2");
                filepath2 = jfc.getSelectedFile().getAbsolutePath();
            }
        });
        tupian.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser jfc = new JFileChooser();
                jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
                jfc.setFileFilter(new FileFilter() {
                    @Override
                    public boolean accept(File file) {
                        String name = file.getName();
                        return file.isDirectory() || name.toLowerCase().endsWith(".png") || name.toLowerCase().endsWith(".jpg");  // 仅显示图片文件
                    }

                    @Override
                    public String getDescription() {
                        return "*.jpg;*.png";
                    }
                });
                jfc.showDialog(new JLabel(), "选择图片");
                picpath = jfc.getSelectedFile().getAbsolutePath();
            }
        });
        open.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser jfc = new JFileChooser();
                jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
                jfc.setFileFilter(new FileFilter() {
                    @Override
                    public boolean accept(File file) {
                        String name = file.getName();
                        return file.isDirectory() || name.toLowerCase().endsWith(".xls") || name.toLowerCase().endsWith(".xlsx");  // 仅显示目录和xls、xlsx文件
                    }

                    @Override
                    public String getDescription() {
                        return "*.xls;*.xlsx";
                    }
                });
                jfc.showDialog(new JLabel(), "选择");
                absolutePath = jfc.getSelectedFile().getAbsolutePath();
            }
        });
        //确定按钮
        submit.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {

                try {


                    String subject = jtfIP.getText();
                    String text = jtfDatabaseName.getText();

                    EmailSender es = new EmailSender();
                    es.setParams(subject,text,filepath1,filepath2,picpath);

                    //判断是不是已经选择了文件夹
                    if (absolutePath.isEmpty()) {
                        JOptionPane.showMessageDialog(jp, "请先选择Excel文件", "错误", JOptionPane.WARNING_MESSAGE);
                    } else {
                        ReadRemark wr = new ReadRemark();
                       // wr.setMysqlconn(mysqlconn);
                        wr.setExcelPath(absolutePath);
                        //wr.setMysqlName(userName);
                        //wr.setMysqlurl(mysqlurl);
                        ///检查Excel的框架是不是正确
                        ReturnMessage checkMessage = wr.checkAndSaveRightExcel();
                        if (checkMessage.isSuccess()) {
                            ReturnMessage writeMessage = wr.write();
                            //判断是不是插入成功
                            if (writeMessage.isSuccess()) {
                                JOptionPane.showMessageDialog(jp, writeMessage.getReturnMessageContent(), "成功", JOptionPane.PLAIN_MESSAGE);
                            }else{
                                JOptionPane.showMessageDialog(jp, writeMessage.getReturnMessageContent(), "失败", JOptionPane.WARNING_MESSAGE);
                            }
                        } else {
                            JOptionPane.showMessageDialog(jp, checkMessage.getReturnMessageContent(),"失败" , JOptionPane.WARNING_MESSAGE);
                        }
                        fujian1.removeAll();
                        fujian2.removeAll();
                        tupian.removeAll();
                        jp.repaint();
                        jp.updateUI();
                    }
                } catch (Exception e2) {
                    e2.printStackTrace();
                }
            }
        });

    }

}
