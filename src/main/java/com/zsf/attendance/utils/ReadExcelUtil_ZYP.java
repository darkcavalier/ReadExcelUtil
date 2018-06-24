package com.zsf.attendance.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

public class ReadExcelUtil_ZYP {
	private static List<String> dateList = new ArrayList<String>();
	private static List<String> timeList = new ArrayList<String>();
	private static List<String> nameList = new ArrayList<String>();
	private static int num;

	// 读取，全部sheet表及数据
	public static void showExcel() throws Exception {
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(new File("D:/员工刷卡记录表.xls")));
		HSSFSheet sheet = null;
		// i表示sheet表格 j控制行，k控制列
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			sheet = workbook.getSheetAt(i);
			for (int j = 0; j < sheet.getLastRowNum() + 1; j++) {
				HSSFRow row = sheet.getRow(j);
				int value = j % 3;
				if (row != null) {
					switch (value) {
					case 0: //工号，姓名
						getCommInfo(row);
						break;
					case 1: //日期
						getDateInfo(row);
						break;
					case 2://时间
						getTimeInfo(row);
						break;
					}
				}
			}
			System.out.println("读取sheet表：" + workbook.getSheetName(i) + " 完成");
		}
	}

	// 姓名 工号 部门信息
	private static void getCommInfo(HSSFRow row) {
		for (int k = 0; k < 19; k++) {
			HSSFCell cell = row.getCell(k);
			if (cell != null) {
				System.out.print(row.getCell(k) + " ");// 表示第j行第k列
			} else {
				System.out.println("");
			}

		}

		nameList.add(row.getCell(10).getStringCellValue());
	}

	// 日期
	private static void getDateInfo(HSSFRow row) {
		for (int k = 0; k < 8; k++) {
			HSSFCell cell = row.getCell(k);
			if (cell != null) {
				String dateResult = String.valueOf(row.getCell(k).getNumericCellValue());
				dateResult = dateResult.substring(0, dateResult.lastIndexOf(".")) + "日：";
				if (dateList.size() < 5) {
					dateList.add(dateResult);
				}
			}
		}
	}

	// 具体时间，工钱
	private static void getTimeInfo(HSSFRow row) {
		float starthour, endhour, startmin, endmin;
		float time = 0;// 加班时长
		int day = 0;// 加班满2.5小时天数
		int latetimes = 0, learlytimes = 0;
		String sendhour;
		System.out.println();
		for (int k = 0; k < 5; k++) {
			HSSFCell cell = row.getCell(k);
			if (cell != null) {
				String dateTime = String.valueOf(row.getCell(k).getStringCellValue());
				dateTime = dateTime.replaceAll("\r|\n", "-");
				if ("".equals(dateTime)) {
					timeList.add("无记录");
				} else {
					timeList.add(dateTime);
				}

			}
		}

		for (int i = 0; i < dateList.size(); i++) {
			System.out.println(dateList.get(i) + timeList.get(i));
			float t = 0;
			if (timeList.get(i).length() >= 11) {
				sendhour = timeList.get(i).substring(timeList.get(i).length() - 5, timeList.get(i).length() - 3);
				if (sendhour.matches("[0-9]+")) {
					starthour = Float.parseFloat(timeList.get(i).substring(0, 2));
					endhour = Float.parseFloat(sendhour);
					startmin = Float.parseFloat(timeList.get(i).substring(3, 5));
					endmin = Float.parseFloat(
							timeList.get(i).substring(timeList.get(i).length() - 2, timeList.get(i).length()));

					if (starthour == 9 && startmin <= 30 && startmin > 0) {
						if (endhour > 18) {

							t = endhour - 19 + convert(endmin / 60);
						} else if (endhour < 18) {
							// 早退
							learlytimes++;
						}
					} else if ((starthour == 9 && startmin == 0) || (starthour < 9)) {
						if (endhour > 18 && endmin >= 30) {

							t = endhour - 18 + convert((endmin - 30) / 60);
						} else if (endhour > 19 && endmin < 30) {

							t = endhour - 19 + convert((30 - endmin) / 60);
						} else if ((endhour == 19 && endhour < 30)) {
							t = (float) 0.5;
						} else if ((endhour == 17 && endmin < 30) || (endhour < 17)) {
							// 早退
							learlytimes++;
						}

					} else {
						// 迟到
						latetimes++;
					}
				}
				if (t >= 2.5) {
					day++;
				}
				time += t;
				// time+=(float)Math.round(t*10)/10;
			}

			// System.out.println((float)Math.round(t*10)/10);
		}
		System.out.println(nameList.get(num) + " " + "加班总时长：" + time + "小时，" + "迟到次数：" + latetimes + "次，" + "早退次数："
				+ learlytimes + "次，" + "加班费：" + (day * 20) + "元");
		num++;
		dateList.clear();
		timeList.clear();
	}

	// 分钟换算
	public static float convert(float daymin) {
		if (daymin >= 0.5) {
			return (float) 0.5;
		} else {
			return 0;
		}
	}

	// 读取，指定sheet表及数据
	public void showExcel2() throws Exception {
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(new File("E:/temp/t1.xls")));
		HSSFSheet sheet = null;
		int i = workbook.getSheetIndex("xt"); // sheet表名
		sheet = workbook.getSheetAt(i);
		for (int j = 0; j < sheet.getLastRowNum() + 1; j++) {// getLastRowNum
																// 获取最后一行的行标
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
		try {
			showExcel();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}