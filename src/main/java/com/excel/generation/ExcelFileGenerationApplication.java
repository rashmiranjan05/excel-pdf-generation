package com.excel.generation;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ExcelFileGenerationApplication {

	public static void main(String[] args) {
		SpringApplication.run(ExcelFileGenerationApplication.class, args);
	}

//	public void generateExcelFile(){
//		HSSFWorkbook workbook = new HSSFWorkbook();
//		HSSFSheet sheet = workbook.createSheet("Bill payment report");
//	}

}
