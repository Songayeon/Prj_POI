package com.study.prj_poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadAndWriteExcelProgram {

	public static void main(String... args) {
		new ReadAndWriteExcelProgram();
	}

	public ReadAndWriteExcelProgram() {
		
		String version = "xls";
//		String version = "xlsx";
		
		// Workbook 취득
		Workbook workbook = getWorkbook("C:\\Dev_Info\\work\\SampleTest.xls", version);
		
		
		// Workbook 안에 시트 취득
		Sheet sheet = workbook.getSheetAt(0);
		
		
		// Sheet 에서 셀 취득 후 데이터 설정
		// 1월 오전 
		getCell(sheet, 1, 2).setCellValue(-5);
		// 1월 오후 
		getCell(sheet, 1, 3).setCellValue(0);
		
		// 2월 오전 
		getCell(sheet, 2, 2).setCellValue(-10);
		// 2월 오후 
		getCell(sheet, 2, 3).setCellValue(-5);
		
		// 3월 오전 
		getCell(sheet, 3, 2).setCellValue(0);
		// 3월 오후 
		getCell(sheet, 3, 3).setCellValue(2);
		
		// 4월 오전 
		getCell(sheet, 4, 2).setCellValue(4);
		// 4월 오후 
		getCell(sheet, 4, 3).setCellValue(10);
		
		// 5월 오전 
		getCell(sheet, 5, 2).setCellValue(10);
		// 5월 오후 
		getCell(sheet, 5, 3).setCellValue(15);
		
		// 6월 오전 
		getCell(sheet, 6, 2).setCellValue(18);
		// 6월 오후 
		getCell(sheet, 6, 3).setCellValue(25);
		
		// 7월 오전 
		getCell(sheet, 7, 2).setCellValue(23);
		// 7월 오후 
		getCell(sheet, 7, 3).setCellValue(28);
		
		// 8월 오전 
		getCell(sheet, 8, 2).setCellValue(25);
		// 8월 오후 
		getCell(sheet, 8, 3).setCellValue(31);
		
		// 9월 오전 
		getCell(sheet, 9, 2).setCellValue(25);
		// 9월 오후 
		getCell(sheet, 9, 3).setCellValue(29);
		
		// 10월 오전 
		getCell(sheet, 10, 2).setCellValue(15);
		// 10월 오후 
		getCell(sheet, 10, 3).setCellValue(25);
		
		// 11월 오전 
		getCell(sheet, 11, 2).setCellValue(11);
		// 11월 오후 
		getCell(sheet, 11, 3).setCellValue(17);
		
		// 12월 오전 
		getCell(sheet, 12, 2).setCellValue(5);
		// 12월 오후 
		getCell(sheet, 12, 3).setCellValue(9);
		
		
		// 함수값 재설정
		for(int i=1; i<=12; i++) {
			getCell(sheet,  i, 1).setCellFormula(String.format("AVERAGE(C%d:D%d)", i+1, i+1));
		}
		writeExcel(workbook, "C:\\Dev_Info\\work\\SampleTestUpdate." + version);
	}
	
	
	/**
	 * Workbook 읽어 드리기
	 * @param filename
	 * @param version
	 * @return
	 */
	public Workbook getWorkbook(String filename, String version) {
		try(FileInputStream stream = new FileInputStream(filename)){
			
			// 표준 xml 버전
			if("xls".equalsIgnoreCase(version)) {
				return new HSSFWorkbook(stream);
				
			// 확장 xlsx 버전
			}else if("xlsx".equalsIgnoreCase(version)) {
				return new XSSFWorkbook(stream);
			}
			throw new NoClassDefFoundError();
			
		}catch (Throwable e) {
			e.printStackTrace();
			return null;
		}
	}
	
	/**
	 * Sheet 로 부터 Row를 취득, 생성하기 
	 * @param sheet
	 * @param rownum
	 * @return
	 */
	public Row getRow(Sheet sheet, int rownum) {
		Row row = sheet.getRow(rownum);
		if(row == null) {
			row = sheet.createRow(rownum);
		}
		return row;
	}
	
	/**
	 * Row 로 부터 Cell을 취득, 생성하기
	 * @param row
	 * @param cellnum
	 * @return
	 */
	public Cell getCell(Row row, int cellnum) {
		Cell cell = row.getCell(cellnum);
		if(cell == null) {
			cell = row.createCell(cellnum);
		}
		return cell;
	}
	
	/**
	 * Sheet 로 부터 Row를 → Cell을 취득, 생성하기 
	 * @param sheet
	 * @param rownum
	 * @param cellnum
	 * @return
	 */
	public Cell getCell(Sheet sheet, int rownum, int cellnum) {
		Row row = getRow(sheet, rownum);
		return getCell(row, cellnum);
	}
	
	/**
	 * 엑셀 파일 쓰기
	 * @param workbook
	 * @param filepath
	 */
	public void writeExcel(Workbook workbook, String filepath) {
		try (FileOutputStream stream = new FileOutputStream(filepath)){
			workbook.write(stream);
		}catch (Throwable e) {
			e.printStackTrace();
		}
	}
}
