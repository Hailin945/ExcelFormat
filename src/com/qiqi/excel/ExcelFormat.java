package com.qiqi.excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelFormat {
	
	// 点这运行
	public static void main(String[] args) {

		List<ExcelDTO> readExcel = readExcel();
		createExcel(readExcel);
	}
	
	
	

	private static List<ExcelDTO> readExcel() {
		List<ExcelDTO> list = new ArrayList<>();
		HSSFWorkbook workbook = null;

		try {
			// 读取Excel文件
			InputStream inputStream = new FileInputStream("/Users/yjj/Desktop/order.xls");
			workbook = new HSSFWorkbook(inputStream);
			inputStream.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

		// 循环工作表
		for (int numSheet = 0; numSheet < workbook.getNumberOfSheets(); numSheet++) {
			HSSFSheet hssfSheet = workbook.getSheetAt(numSheet);
			if (hssfSheet == null) {
				continue;
			}
			// 循环行
			for (int rowNum = 3; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
				HSSFRow hssfRow = hssfSheet.getRow(rowNum);
				if (hssfRow == null) {
					continue;
				}

				// 将单元格中的内容存入集合
				ExcelDTO excelDTO = new ExcelDTO();

				HSSFCell cell = hssfRow.getCell(2);
				if (cell == null) {
					continue;
				}
				excelDTO.setCompanyName(cell.getStringCellValue());

				cell = hssfRow.getCell(17);
				if (cell == null) {
					continue;
				}
				excelDTO.setOrderType(cell.getStringCellValue());


				list.add(excelDTO);
			}
		}
		
		Map<String, List<ExcelDTO>> map = new HashMap<>();
		for (ExcelDTO excelDTO : list) {
			String companyName = excelDTO.getCompanyName();
			if (!map.containsKey(companyName)) {
				map.put(companyName, new ArrayList<>());
			}
			System.out.println(excelDTO.getCompanyName()+"----"+excelDTO.getOrderType());
		}
		
		for (ExcelDTO excelDTO : list) {
			map.get(excelDTO.getCompanyName()).add(excelDTO);
		}
		
		List<ExcelDTO> resultList = new ArrayList<>();
		
		Set<String> keySet = map.keySet();
		for (String name : keySet) {
			List<ExcelDTO> excelDTOs = map.get(name);
			int cancelTotalNum = 0;
			int completedTotalNum = 0;
			int totalNum = excelDTOs.size();
			for (ExcelDTO excelDTO : excelDTOs) {
				if ("已取消".equals(excelDTO.getOrderType())) {
					cancelTotalNum ++;
				}
				if ("已完成".equals(excelDTO.getOrderType())) {
					completedTotalNum ++;
				}
			}
			ExcelDTO result = new ExcelDTO();
			result.setCompanyName(name);
			result.setCancenlTotalNum(cancelTotalNum);
			result.setCompletedTotalNum(completedTotalNum);
			result.setTotalNum(totalNum); 
			resultList.add(result);
		}
		
		
		
//		for (ExcelDTO excel : resultList) {
//			System.out.println(excel.getCompanyName()+"---"+excel.getTotalNum()+"---"+excel.getCancenlTotalNum()+"---"+excel.getCompletedTotalNum());
//		}
		
		return resultList;
	}
	
	 /**
     * 创建Excel
     */
    private static void createExcel(List<ExcelDTO> list) {
    	
        // 创建一个Excel文件
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 创建一个工作表
        HSSFSheet sheet = workbook.createSheet("order");
        // 添加表头行
        HSSFRow hssfRow = sheet.createRow(0);
        // 设置单元格格式居中
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);

        // 添加表头内容
        HSSFCell headCell = hssfRow.createCell(0);
        headCell.setCellValue("公司名称");
        headCell.setCellStyle(cellStyle);

        headCell = hssfRow.createCell(1);
        headCell.setCellValue("总订单数");
        headCell.setCellStyle(cellStyle);

        headCell = hssfRow.createCell(2);
        headCell.setCellValue("已取消");
        headCell.setCellStyle(cellStyle);
        
        headCell = hssfRow.createCell(3);
        headCell.setCellValue("进行中");
        headCell.setCellStyle(cellStyle);
        
        headCell = hssfRow.createCell(4);
        headCell.setCellValue("已完成");
        headCell.setCellStyle(cellStyle);

        // 添加数据内容
        for (int i = 0; i < list.size(); i++) {
            hssfRow = sheet.createRow((int) i + 1);
            ExcelDTO excelDTO = list.get(i);

            // 创建单元格，并设置值
            HSSFCell cell = hssfRow.createCell(0);
            cell.setCellValue(excelDTO.getCompanyName());
            cell.setCellStyle(cellStyle);

            cell = hssfRow.createCell(1);
            cell.setCellValue(excelDTO.getTotalNum());
            cell.setCellStyle(cellStyle);

            cell = hssfRow.createCell(2);
            cell.setCellValue(excelDTO.getCancenlTotalNum());
            cell.setCellStyle(cellStyle);
            
            cell = hssfRow.createCell(3);
            cell.setCellValue(excelDTO.getTotalNum() - excelDTO.getCancenlTotalNum() - excelDTO.getCompletedTotalNum());
            cell.setCellStyle(cellStyle);
            
            cell = hssfRow.createCell(4);
            cell.setCellValue(excelDTO.getCompletedTotalNum());
            cell.setCellStyle(cellStyle);
        }

        // 保存Excel文件
        try {
            OutputStream outputStream = new FileOutputStream("/Users/yjj/Desktop/abc.xls");
            workbook.write(outputStream);
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
   
	
}
