package com.zsf.attendance.utils;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class ReadExcelUtil {

    private static Logger logger = LoggerFactory.getLogger(ReadExcelUtil.class);


    private static int DEFAULT_RECORD_LENGTH = 5; //打卡时间长度 09:30  长度为5
    private static List<String> dateList = new ArrayList<String>();
    private static List<String> timeList = new ArrayList<String>();
    private static final String BLANK = " ";

    // 读取，全部sheet表及数据
    public static void showExcel() {
        long startTime = System.currentTimeMillis();
        FileInputStream inputStream = null;
        try {
            inputStream = new FileInputStream(new File("D:/员工刷卡记录表.xls"));
            HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
            HSSFSheet sheet = null;
            //j控制行，k控制列
            sheet = workbook.getSheetAt(0);
            //获取所有员工的工号
            //Map<String, String> workNumMap = getAllworkNum(sheet);
            for (int j = 0; j < sheet.getLastRowNum(); j++) {
                HSSFRow row = sheet.getRow(j);//第j行对象
                int value = j % 3;
                if (row != null) {
                    switch (value) {
                        case 0:
                            //这里只含有工号那一行的数据 暂时只读取前20列数据
                            getCommInfo(row);
                            break;
                        case 1:
                            // 这里只含有日期那一行的数据  暂时只读取前8列数据，row.getLastCellNum();
                            getDateInfo(row);
                            break;
                        case 2:
                            // 这里只含有时间那一行的数据
                            getTimeInfo(row);
                            break;
                    }
                }
            }
            System.out.println("以下为已离职或者不打卡员工列表");
            getEmployeeMap();
        } catch (IOException e) {
            logger.info("showExcel方法出现异常:message{}" + e.getMessage());
        } finally {
            if (inputStream != null) { //关闭流，释放资源
                try {
                    inputStream.close();
                } catch (IOException e) {
                    logger.info(e.getMessage());
                }
            }
        }
        long endTime = System.currentTimeMillis();
        System.out.println("本次操作一共耗时:" + (endTime - startTime) + "ms");
    }


    //获取所有员工的工号
    private static Map<String, String> getAllworkNum(HSSFSheet sheet) {
        Map<String, String> map = new LinkedHashMap<String, String>();
        for (int i = 0; i < sheet.getLastRowNum(); i = i + 3) {
            HSSFRow row = sheet.getRow(i);
            String str = "";
            HSSFCell cell = row.getCell(2);  //员工号在第三列
            int type = cell.getCellType();
            if (type == 0) {
                str = String.valueOf(cell.getNumericCellValue());
            } else if (type == 1) {
                str = cell.getStringCellValue();
            }
            String result = str.endsWith(".0") ? str.replaceAll(".0", "") : str;
            map.put(result, result);
        }
        return map;
    }


    //这里只含有工号那一行的数据 暂时只读取前20列数据
    private static void getCommInfo(HSSFRow row) {
        String str = "";
        for (int k = 0; k < 19; k++) {
            HSSFCell cell = row.getCell(k);
            if (cell != null) {
                int type = cell.getCellType();
                if (type == 0) {
                    str += String.valueOf(cell.getNumericCellValue()) + BLANK;
                    str=str.contains(".0")?str.replace(".0",BLANK):str;
                } else if (type == 1) {
                    str += cell.getStringCellValue().contains(".0") ? cell.getStringCellValue().replaceAll(".0", BLANK) : cell.getStringCellValue();
                }
            } else {
                System.out.println("");
            }
        }
        System.out.print(str);
    }


    // 这里只含有日期那一行的数据  暂时只读取前8列数据，row.getLastCellNum();
    private static void getDateInfo(HSSFRow row) {
        for (int k = 0; k < 8; k++) {
            HSSFCell cell = row.getCell(k);
            if (cell != null) {
                String dateResult = String.valueOf(row.getCell(k).getNumericCellValue());
                dateResult = dateResult.substring(0, dateResult.lastIndexOf(".")) + "日:";
                if (dateList.size() < 5) {
                    dateList.add(dateResult);
                }
            }
        }
    }


    // 这里只含有时间那一行的数据
    private static void getTimeInfo(HSSFRow row) {
        System.out.println();
        for (int k = 0; k < 5; k++) {
            HSSFCell cell = row.getCell(k);
            if (cell != null) {
                String result = row.getCell(k).getStringCellValue().replaceAll("\n", "-");
                if ("".equals(result)) {
                    timeList.add("未打卡");
                } else {
                    timeList.add(result);
                }
            }
        }
        for (int i = 0; i < dateList.size(); i++) {
            System.out.println(dateList.get(i) + timeList.get(i));
        }
        dateList.clear();
        timeList.clear();
    }


    private static void getEmployeeMap() {
        //定义一个已离职或者不打卡员工map
        Map<String, String> map = new HashMap<String, String>();
        map.put("100", "俞耀祖");
        map.put("107", "曹小波");
        map.put("131", "林文生");
        map.put("347", "叶文斌");
        map.put("850", "包蕾");
        map.put("860", "陈天成");
        map.put("1284", "贺鹏");
        map.put("1411", "李超");
        map.put("1626", "吴俊华");
        map.put("2207", "候晓燕");
        map.put("2218", "杨尚剑");
        map.put("2253", "朱宝龙");
        map.put("2347", "王丽敏");
        map.put("2373", "刘少磊");
        map.put("2410", "杨永鹏");
        map.put("2974", "王鹏");
        map.put("3006", "孙志福");
        map.put("3062", "谢磊");
        map.put("3065", "朱佩佩");
        map.put("3886", "赵肖肖");
        map.put("8001", "马学洋");
        map.put("8002", "任鹏冲");
        map.put("8004", "李圆圆");
        map.put("8005", "吉宜洛");
        map.put("8006", "胡亚熙");
        map.put("9003", "孙峥嵘");
        System.out.println();
        if (!map.isEmpty() && map.size() > 0) {
            for (String key : map.keySet()) {
                System.out.println("工号:" + key + "  姓名:" + map.get(key));
            }
        }
    }

    // 读取，指定sheet表及数据
    public void showExcel2() throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(new File("E:/temp/t1.xls")));
        HSSFSheet sheet = null;
        int i = workbook.getSheetIndex("xt"); // sheet表名
        sheet = workbook.getSheetAt(i);
        for (int j = 0; j < sheet.getLastRowNum() + 1; j++) {// getLastRowNum
            HSSFRow row = sheet.getRow(j);
            if (row != null) {
                for (int k = 0; k < row.getLastCellNum(); k++) {// getLastCellNum
                    // 是获取最后一个不为空的列是第几个
                    if (row.getCell(k) != null) { // getCell 获取单元格数据
                        System.out.print(row.getCell(k) + "\t");
                    } else {
                        System.out.print("\t");
                    }
                }
            }
            System.out.println("");
        }
    }

    // 写入，往指定sheet表的单元格
    public void insertExcel3() throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(new File("E:/temp/t1.xls"))); // 读取的文件
        HSSFSheet sheet = null;
        int i = workbook.getSheetIndex("xt"); // sheet表名
        sheet = workbook.getSheetAt(i);

        HSSFRow row = sheet.getRow(0); // 获取指定的行对象，无数据则为空，需要创建
        if (row == null) {
            row = sheet.createRow(0); // 该行无数据，创建行对象
        }
        Cell cell = row.createCell(1); // 创建指定单元格对象。如本身有数据会替换掉
        cell.setCellValue("tt"); // 设置内容
        FileOutputStream fo = new FileOutputStream("E:/temp/t1.xls"); // 输出到文件
        workbook.write(fo);

    }


    public static void main(String[] args) {
        showExcel();
    }
}
