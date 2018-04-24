
import java.io.*;
import java.sql.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadRemark {

    //定义常量
    private String excelPath = "e:\\tables.xlsx";
    private Workbook wb = null;
    //private Connection mysqlconn=null;
    //检查Excel的格式是否正确
    public ReturnMessage checkAndSaveRightExcel() {
        ReturnMessage message = new ReturnMessage(true);
        try {
            //所有表必须按照格式写否则报错
            InputStream is = new FileInputStream(excelPath);
            // 处理xlsx文件
            if (excelPath.endsWith("xlsx")) {
                wb = new XSSFWorkbook(is);
                //判断文件打开是否成功

            } else if (excelPath.endsWith("xls")) {
                // 处理xls文件
                wb = new HSSFWorkbook(is);
            }

            //检查文件Excel内有没有sheet
            int numberOfSheets = wb.getNumberOfSheets();
            if (numberOfSheets == 0) {
                message.setSuccess(false);
                message.setReturnMessageContent("Excel文件内没有表");
                return message;
            }
        } catch (Exception e) {
            message.setSuccess(false);
            message.setReturnMessageContent("发生未知错误");
        }
        return message;
    }

    //发送邮件的主要方法
    public ReturnMessage write() {

        ReturnMessage message = new ReturnMessage(true);
        HashMap<String, String> map = null;
        ArrayList<String> mails1 = null;
        ArrayList<String> mails2 = null;
        ArrayList<Long> al = new ArrayList();
        //Statement stmt=null;
        //连接数据库
        try {

            //for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet1 = wb.getSheetAt(0);
            Sheet sheet2 = wb.getSheetAt(1);
            //map = readColumnAndRemarkFromExcel(sheet1);
            mails1 = readMail(sheet1);
            mails2 = readMail(sheet2);
            int count1 = 0;
            EmailSender es = new EmailSender();
            for (String recievemail : mails1){
                es.sendMail(mails2.get(count1/2),recievemail);
                count1++;
            }
            boolean allMatch = true;

            if (allMatch) {
                message.setSuccess(true);
                message.setReturnMessageContent("导入成功");
                count1 = 0;
            } else {
                message.setSuccess(true);
                message.setReturnMessageContent("导入成功，但sheet内容与对应的表的字段名称不完全匹配,建议重新检查表名");
            }

            //}

        } catch (Exception e){
            e.printStackTrace();
        }
        return message;
    }

    private HashMap<String, String> readColumnAndRemarkFromExcel(Sheet sheet) {
        HashMap<String, String> result = new HashMap();

        String sheetName = sheet.getSheetName();
        // 得到总行数
        int rowNum = sheet.getLastRowNum();
        // 读取正文,只读取第一列,如果某一字段有空值，或者出现中文,返回空的list.
        for (int i = 1; i <= rowNum; i++) {
            Row row = sheet.getRow(i);
            String columnStr = getCellFormatValue(row.getCell((short) 1));
            // 判断第一列是否有空值
            if (columnStr.trim().equals("")) {
                continue;
            }

            String remarkStr = getCellFormatValue(row.getCell((short) 2));
            // 判断第二列是否有空值
            if (remarkStr.trim().equals("")) {
                continue;
            }
            result.put(columnStr, remarkStr);
        }


        return result;
    }

   //读取sheet中的邮箱list
    private ArrayList<String> readMail(Sheet sheet) {
        ArrayList<String> result = new ArrayList<String>();

        String sheetName = sheet.getSheetName();
        // 得到总行数
        int rowNum = sheet.getLastRowNum();
        // 读取正文,只读取第一列,如果某一字段有空值，或者出现中文,返回空的list.
        for (int i = 1; i <= rowNum; i++) {
            Row row = sheet.getRow(i);
            String mailName = getCellFormatValue(row.getCell((short) 0));
            // 判断第一列是否有空值
            if (mailName.trim().equals("")) {
                continue;
            }
            result.add(mailName);
        }
        return result;
    }

   //检验邮箱格式是不是正确
    private boolean checkEmail(String email){
        boolean flag = false;
        try{
            String check = "^([a-z0-9A-Z]+[-|\\.]?)+[a-z0-9A-Z]@([a-z0-9A-Z]+(-[a-z0-9A-Z]+)?\\.)+[a-zA-Z]{2,}$";
            Pattern regex = Pattern.compile(check);
            Matcher matcher = regex.matcher(email);
            flag = matcher.matches();
        }catch(Exception e){
            flag = false;
        }
        return flag;
    }

    /**
     * 根据HSSFCell类型设置数据
     *
     * @param cell
     * @return
     */

    private String getCellFormatValue(Cell cell) {
        String cellvalue = "";
        if (cell != null) {
            // 判断当前Cell的Type
            switch (cell.getCellType()) {
                // 如果当前Cell的Type为NUMERIC
                case HSSFCell.CELL_TYPE_NUMERIC:
                    cellvalue = String.valueOf(cell.getNumericCellValue());
                    break;
                // 如果当前Cell的Type为STRIN
                case HSSFCell.CELL_TYPE_STRING:
                    // 取得当前的Cell字符串
                    cellvalue = cell.getRichStringCellValue().getString();
                    break;
                case HSSFCell.CELL_TYPE_BOOLEAN:
                    cellvalue = String.valueOf(cell.getBooleanCellValue());
                    break;
                // 默认的Cell值
                default:
                    cellvalue = " ";
            }
        } else {
            cellvalue = "";
        }
        return cellvalue;

    }

    /*public String getMysqlurl() {
        return mysqlurl;
    }

    public void setMysqlurl(String mysqlurl) {
        this.mysqlurl = mysqlurl;
    }*/

    public String getExcelPath() {
        return excelPath;
    }

    public void setExcelPath(String excelPath) {
        this.excelPath = excelPath;
    }

    /*public String getMysqlName() {
        return mysqlName;
    }

    public void setMysqlName(String mysqlName) {
        this.mysqlName = mysqlName;
    }

    public String getMysqlPassword() {
        return mysqlPassword;
    }

    public void setMysqlPassword(String mysqlPassword) {
        this.mysqlPassword = mysqlPassword;
    }*/

    /*public Connection getMysqlconn() {
        return mysqlconn;
    }

    public void setMysqlconn(Connection mysqlconn) {
        this.mysqlconn = mysqlconn;
    }*/
}

