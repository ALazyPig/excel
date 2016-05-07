package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ReadExcelPoi {

	public static void main(String[] args) {
		//��Ҫ������excel�ļ�
		File file = new File("C:\\Users\\333666666\\Desktop\\test.xls");
		FileInputStream fileInputStream = null;
		try {
			fileInputStream = new FileInputStream(file);
			HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
			//��ȡ��һ�����������ַ�ʽ
			//HSSFSheet sheet = workbook.getSheet("Sheet0");
			//��ȡĬ�ϵĵ�һ��sheetҳ
			HSSFSheet sheet = workbook.getSheetAt(0);
			
			int firstRowNum = sheet.getFirstRowNum();
			//System.out.println("" + firstRowNum);
			//��ȡsheet�����һ���к�
			int lastRowNum = sheet.getLastRowNum();
			
			for(int i = 0; i < lastRowNum; i++){
				HSSFRow row = sheet.getRow(i);
				
				//��ȡ��ǰ�����һ����Ԫ���к�
				int lastCellNum = row.getLastCellNum();
				for(int j = 0; j < lastCellNum; j++){
					HSSFCell cell = row.getCell(j);
					String value = cell.getStringCellValue();
					System.out.print(value + " ");
				}
				System.out.println();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		//
		
		
	}

}
