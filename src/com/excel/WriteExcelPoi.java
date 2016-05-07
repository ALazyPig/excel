package com.excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class WriteExcelPoi {
	public static void main(String[] args){
		//�����ͷ
		String[] title = {"id","name","sex"};
		//����Excel������
		HSSFWorkbook workbook = new HSSFWorkbook();
		//����������
		HSSFSheet sheet = workbook.createSheet();
		//������һ��
		HSSFRow row = sheet.createRow(0);
		//���嵥Ԫ��
		HSSFCell cell = null;
		//׷������
		for(int i = 0; i < title.length; i++){
			cell = row.createCell(i);
			cell.setCellValue(title[i]);
		}
		//׷������
		for(int i = 1; i < 10; i++){
			HSSFRow nextrow = sheet.createRow(i);
			HSSFCell cell2 = nextrow.createCell(0);
			cell2.setCellValue("a" + i);
			cell2 = nextrow.createCell(1);
			cell2.setCellValue("user" + i);
			cell2 = nextrow.createCell(2);
			cell2.setCellValue("male");
		}
		//����excel
		File file = new File("C:\\Users\\333666666\\Desktop\\test.xls");
		FileOutputStream stream = null;
		try {
			if(!file.exists())	file.createNewFile();
			String name = file.getName();
			String path = file.getPath(); 
			System.out.println("name:" + name + "path:" + path);
		} catch (IOException e1) {
			e1.printStackTrace();
		}
		try {
			stream = new FileOutputStream(file);
			workbook.write(stream);
		} catch (Exception e) {
			e.printStackTrace();
		}finally{
			try {
				if(stream != null) stream.close();
				if(workbook != null) workbook.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
			
		}
	}
}
