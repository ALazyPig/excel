package com.excel.template;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.jdom.Attribute;
import org.jdom.Document;
import org.jdom.Element;
import org.jdom.JDOMException;
import org.jdom.input.SAXBuilder;


public class CreateTemplate {

	public static void main(String[] args) {
		String path = System.getProperty("user.dir")+"\\bin\\student.xml";
		System.out.println(path);
		File file = new File(path);
		SAXBuilder builder = new SAXBuilder();
		try {
			//解析xml
			Document parse = builder.build(file);
			
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFSheet sheet = workbook.createSheet("Sheet0");
			//获取根节点
			Element root = parse.getRootElement();
			//获取模板名称  name="学生信息导入"
			String templateName = root.getAttribute("name").getValue();
			
			int rownum = 0;
			int column = 0;
			
			//设置列宽
			Element colgroup = root.getChild("colgroup");
			new CreateTemplate().setColumnWidth(sheet,colgroup);
			
			//设置标题
			Element title = root.getChild("title");
			List<Element> trs = title.getChildren("tr");
			for (int i = 0; i < trs.size(); i++) {
				Element tr = trs.get(i);
				List<Element> tds = tr.getChildren("td");
				HSSFRow row = sheet.createRow(rownum);
				HSSFCellStyle cellStyle = workbook.createCellStyle();
				cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
				for(column = 0;column <tds.size();column ++){
					Element td = tds.get(column);
					HSSFCell cell = row.createCell(column);
					Attribute rowSpan = td.getAttribute("rowspan");
					Attribute colSpan = td.getAttribute("colspan");
					Attribute value = td.getAttribute("value");
					if(value !=null){
						String val = value.getValue();
						cell.setCellValue(val);
						//excel行列从0开始
						int rspan = rowSpan.getIntValue() - 1;
						int cspan = colSpan.getIntValue() -1;
						//设置单元格字体
						HSSFFont font = workbook.createFont();
						font.setFontName("仿宋_GB2312");
						font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//字体加粗
						//font.setFontHeight((short)12);
						font.setFontHeightInPoints((short)12);//高度
						cellStyle.setFont(font);
						cell.setCellStyle(cellStyle);
						//合并单元格居中
						sheet.addMergedRegion(new CellRangeAddress(rspan, rspan, 0, cspan));
					}
				}
				//下次取从标题栏下方开始
				rownum ++;
			}
			//设置表头
			Element thead = root.getChild("thead");
			trs = thead.getChildren("tr");
			for (int i = 0; i < trs.size(); i++) {
				Element tr = trs.get(i);
				HSSFRow row = sheet.createRow(rownum);
				List<Element> ths = tr.getChildren("th");
				for(column = 0;column < ths.size();column++){
					Element th = ths.get(column);
					Attribute valueAttr = th.getAttribute("value");
					HSSFCell cell = row.createCell(column);
					if(valueAttr != null){
						String value =valueAttr.getValue();
						cell.setCellValue(value);
					}
				}
				rownum++;
			}
			//设置数据区域的样式
			Element tbody = root.getChild("tbody");
			Element tr = tbody.getChild("tr");
			int repeat = tr.getAttribute("repeat").getIntValue();
			
			List<Element> tds = tr.getChildren("td");
			for (int i = 0; i < repeat; i++) {
				HSSFRow row = sheet.createRow(rownum);
				for(column =0 ;column < tds.size();column++){
					Element td = tds.get(column);
					HSSFCell cell = row.createCell(column);
					setType(workbook,cell,td);
				}
				rownum++;
			}
			
			File tempFile = new File("e:/" + templateName + ".xls");
			tempFile.delete();
			tempFile.createNewFile();
			FileOutputStream stream = FileUtils.openOutputStream(tempFile);
			workbook.write(stream);
			stream.close();
		} catch (JDOMException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}
	//设置数据区域的样式
	private static void setType(HSSFWorkbook workbook, HSSFCell cell, Element td) {
		Attribute typeAttr = td.getAttribute("type");
		String type = typeAttr.getValue();
		HSSFDataFormat format = workbook.createDataFormat();
		HSSFCellStyle cellStyle = workbook.createCellStyle();
		if("NUMERIC".equalsIgnoreCase(type)){    //数字类型
			cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
			Attribute formatAttr = td.getAttribute("format");
			String formatValue = formatAttr.getValue();
			formatValue = StringUtils.isNotBlank(formatValue)? formatValue : "#,##0.00";
			cellStyle.setDataFormat(format.getFormat(formatValue));
		}else if("STRING".equalsIgnoreCase(type)){    //String类型
			cell.setCellValue("");
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			cellStyle.setDataFormat(format.getFormat("@"));
		}else if("DATE".equalsIgnoreCase(type)){     //日期类型
			cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
			cellStyle.setDataFormat(format.getFormat("yyyy-m-d"));
		}else if("ENUM".equalsIgnoreCase(type)){    //枚举类型
			CellRangeAddressList regions = 
				new CellRangeAddressList(cell.getRowIndex(), cell.getRowIndex(), 
						cell.getColumnIndex(), cell.getColumnIndex());
			Attribute enumAttr = td.getAttribute("format");
			String enumValue = enumAttr.getValue();
			//加载下拉列表内容
			DVConstraint constraint = 
				DVConstraint.createExplicitListConstraint(enumValue.split(","));
			//数据有效性对象
			HSSFDataValidation dataValidation = new HSSFDataValidation(regions, constraint);
			workbook.getSheetAt(0).addValidationData(dataValidation);
		}
		cell.setCellStyle(cellStyle);
	}

	public void setColumnWidth(HSSFSheet sheet, Element colgroup) {
		List<Element> cols = colgroup.getChildren("col");
		for(int i = 0; i < cols.size(); i++){
			Element col = cols.get(i);
			Attribute width = col.getAttribute("width");
			//截取单位，一般为em或px
			String unit = width.getValue().replaceAll("[0-9,\\.]", "");
			//截取值
			String value = width.getValue().replaceAll(unit, "");
			int v=0;
			//换算成excel单位
			if(unit == null || unit == " " || "px".endsWith(unit)){
				v = Math.round(Float.parseFloat(value) * 37F);
			}else if ("em".endsWith(unit)){
				v = Math.round(Float.parseFloat(value) * 267.5F);
			}
			sheet.setColumnWidth(i, v);
		}
	}

}
