package com.seoul.elk.Main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class XlsLoad {

	public static void ExcelParser() {
		
		try {
			File dirFile = new File("/home/spinut/Documents/elk/ES_Data/Dic_kor/");
			File[] fileList = dirFile.listFiles();
			FileInputStream inputStream = null;
			for(File tempFile: fileList) {
				if(tempFile.isFile()) {
					String fileName = tempFile.getName();		// 파일명
					
					// 엑셀파일 아니면 건너뜀
					if(!(fileName.endsWith(".xls"))) {
						continue;
					}
					
					// 파일 Read
					inputStream = new FileInputStream(tempFile.getAbsolutePath());
					HSSFWorkbook workbook = new HSSFWorkbook(inputStream);				// 엑셀읽기
					HSSFSheet sheet = workbook.getSheetAt(0);							// 엑셀문서 첫번째 sheet
					
					int rows = sheet.getPhysicalNumberOfRows();
					HSSFRow row = sheet.getRow(0);
					int cols = row.getLastCellNum();
					System.out.println("Test Row count : [" + rows + "]");
					System.out.println("Test Col count : [" + cols + "]");
					// 0:어휘, 1:단어,  2:고유어, 10:품사(명사),  15:뜻풀이, 18:전문분야 (체육,의학등)
					
					// Row
					for(int i=1; i<2; i++) {
						// column
						for(int j=0; j<cols; j++) {
							HSSFRow curRow = sheet.getRow(i);
							String data = curRow.getCell(j).getStringCellValue();
							System.out.print("data : "+data);
						}
					}
					
					
				}
			}
		}catch(FileNotFoundException fnfe) {
		
		}catch(IOException ioe) {
			
		}finally {
			
		}
	}
}
