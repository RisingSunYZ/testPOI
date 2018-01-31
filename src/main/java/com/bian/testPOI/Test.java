package com.bian.testPOI;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * 
 * @description POI 报表 实现功能 将 数据写入Excel 并上传到本地相应位置
 * @author yz
 * @data 2017年12月12日
 */
public class Test {

	private static List<Map<String,String>> dataSources = null;//定义数据源
	private static final String  path = "D:/test.xls";//定义数据源
	
	static{// 创建模拟数据 map为一行 每个key 对应每个field 
		dataSources = new ArrayList<Map<String,String>>();
		Map<String,String> dataSource = new HashMap<String, String>();
		dataSource.put("name", "yz");
		dataSource.put("addr", "gaoxin");
		dataSource.put("score", "89.2");
		dataSources.add(dataSource);
		Map<String,String> dataSource1 = new HashMap<String, String>();
		dataSource1.put("name", "zwl");
		dataSource1.put("addr", "jingkai");
		dataSource1.put("score", "73");
		dataSources.add(dataSource1);
	}
	
	/**
	 * 读取数据 并导入excel
	 */
	@org.junit.Test
	public void testWritePOI(){
		HSSFWorkbook workBook = new HSSFWorkbook();
		HSSFCellStyle style = workBook.createCellStyle();//创建样式
		HSSFFont font = workBook.createFont();//创建字体
		font.setFontName("宋体");
		font.setFontHeightInPoints((short)14);//设置大小
		style.setFont(font);
		style.setWrapText(true);//设置自动换行 
		style.setBorderLeft(CellStyle.BORDER_THIN);//设置边框
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setAlignment(CellStyle.ALIGN_CENTER);//设置上下居中
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		
		HSSFDataFormat format = workBook.createDataFormat();//处理数字显示格式
		style.setDataFormat(format.getFormat("0.0"));
		
		HSSFSheet sheet = workBook.createSheet("testPOI");//创建一个名为testPOI的工作簿
		int index = 0;
		sheet.setColumnWidth(index++, 3500);//设置列宽 下标从0 开始
		sheet.setColumnWidth(index++, 3500);
		sheet.setColumnWidth(index++, 3500);
		
		HSSFRow titleRow = sheet.createRow(0);//创建标题行
		HSSFCell titleCell = titleRow.createCell(0);//创建单元格
		titleCell.setCellStyle(style);//设置样式
		titleCell.setCellValue("标题");
		titleRow.setHeight((short)800);//设置行高 如果不设置此属性 自动由内容撑开高度
		CellRangeAddress address = new CellRangeAddress(0, 0, 0, index-1);//合并单元格 参数依次为 行start end 列start end
		sheet.addMergedRegion(address);
		
		HSSFRow row = sheet.createRow(1);
		for(int i=0;i<index;i++){
			HSSFCell cell = row.createCell(i);
			cell.setCellStyle(style);
			if(i==0){
				cell.setCellValue("姓名");
			}else if(i == 1){
				cell.setCellValue("地址");				
			}else if(i == 2){
				cell.setCellValue("分数");								
			}else{
				cell.setCellValue("");												
			}
		}
		
		for(int i=0;i<dataSources.size();i++){//读取数据
			HSSFRow rowData = sheet.createRow(i+2);
			for(int j=0;j<index;j++){
				HSSFCell cell = rowData.createCell(j);
				cell.setCellStyle(style);
				if(j==0){
					cell.setCellValue(dataSources.get(i).get("name"));
				}else if(j == 1){
					cell.setCellValue(dataSources.get(i).get("addr"));				
				}else if(j == 2){//数字类型 需要转换 否则在excel 左上角有三角号
					double score = 0.0f;
					if(null!=dataSources.get(i).get("score")&&!dataSources.get(i).get("score").equals("")){
						score = Double.parseDouble(dataSources.get(i).get("score"));						
					}
					cell.setCellValue(score);								
				}else{
					cell.setCellValue("");												
				}
			}
		}
		
		try {
			FileOutputStream fos = new FileOutputStream(path);//导出到盘符 Path 可以在配置文件中配置 name 可以使用uuid
			workBook.write(fos);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * 读取excel数据 存入数据库
	 */
	@org.junit.Test
	public void readPOI(){
		
		try {
			Workbook workBook = WorkbookFactory.create(new FileInputStream(path));
			Sheet sheet = workBook.getSheetAt(0);
			int rowNum = sheet.getLastRowNum();//获取总行数
			for(int row=0;row<=rowNum;row++){
				Row rowTemp = sheet.getRow(row);
				int cellNum = rowTemp.getLastCellNum();//获取每一行的列数
				for(int cellIndex=0;cellIndex<cellNum;cellIndex++){
					Cell cell = rowTemp.getCell(cellIndex);
					switch (cell.getCellType()) {//判断类型 读取数据 double 特殊处理
					case HSSFCell.CELL_TYPE_STRING:
						System.out.println(cell.getStringCellValue());
						break;
					case HSSFCell.CELL_TYPE_NUMERIC:
						System.out.println(cell.getNumericCellValue());
						break;
					case HSSFCell.CELL_TYPE_FORMULA:
						FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
						evaluator.evaluateFormulaCell(cell);
						CellValue cellValue = evaluator.evaluate(cell);
						System.out.println(cellValue.getNumberValue());
						break;
					}
				}
			}
		}  catch (Exception e) {
			e.printStackTrace();
		}
	}
}
