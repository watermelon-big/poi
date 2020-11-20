package com.watermelon;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author cuilei
 * @version 1.0
 * @date 2020/11/20 17:25
 */
public class ExcelWriteTest {
	String  PATH = "E:\\idea-workspace\\POI\\watermelon-pio\\src\\main\\java\\com\\watermelon\\template\\";

	
	//03和07的工作簿类和文件后缀不一样
	
	@Test
	public  void testWrite03() throws IOException {
		//创建一个工作簿
//		Workbook workbook =new HSSFWorkbook();//03
		Workbook workbook = new XSSFWorkbook();//07
		//创建一个工作表
		Sheet sheet = workbook.createSheet("测试表");
		//创建一个行
		Row row1 = sheet.createRow(0);
		//创建单元格
		Cell cell11 = row1.createCell(0);//1,1
		Cell cell12 = row1.createCell(1);//1,2
		
		cell11.setCellValue("今日新增观众");
		cell12.setCellValue(666);

		Row row2 = sheet.createRow(1);
		Cell cell21 = row2.createCell(0);//1,1
		Cell cell22 = row2.createCell(1);//1,2

		cell21.setCellValue("统计时间");
		
		String s = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
		cell22.setCellValue(s);
		
		//生成一张表
		FileOutputStream outputStream = new FileOutputStream(PATH + "观众统计表07.xlsx");
		
		workbook.write(outputStream);
		
		//关闭流
		outputStream.close();

		System.out.println("Excel生成完毕");

	}
}
