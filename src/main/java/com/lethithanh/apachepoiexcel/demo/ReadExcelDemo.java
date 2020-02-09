package com.lethithanh.apachepoiexcel.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
 
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
 
public class ReadExcelDemo {
 
   public static void main(String[] args) throws IOException {
  
       // Doc 1 file XLSX
       FileInputStream inputStream = new FileInputStream(new File("C:/demo/employee.xlsx"));
  
       // Doi tuong workbook cho file XLSX
       XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
 
  
       // Lay ra sheet dau tien tu workbook
       XSSFSheet sheet = workbook.getSheetAt(0);
 
  
       // Lay ra Iterator cho tat ca ca dong cua sheet hien tai
       Iterator<Row> rowIterator = sheet.iterator();
 
       while (rowIterator.hasNext()) {
           Row row = rowIterator.next();
     
           // Lay Iterator cho tat ca cac cell cua dong hien tai
           Iterator<Cell> cellIterator = row.cellIterator();
 
           while (cellIterator.hasNext()) {
               Cell cell = cellIterator.next();
  
               // Doi thanh getCellType() neu su dung POI 4.x
               CellType cellType = cell.getCellTypeEnum();
 
               switch (cellType) {
               case _NONE:
                   System.out.print("");
                   System.out.print("\t");
                   break;
               case BOOLEAN:
                   System.out.print(cell.getBooleanCellValue());
                   System.out.print("\t");
                   break;
               case BLANK:
                   System.out.print("");
                   System.out.print("\t");
                   break;
               case FORMULA:
       
                   // Cong thuc
                   System.out.print(cell.getCellFormula());
                   System.out.print("\t");
                    
                   FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
         
                   // In ra gia tri tu cong thuc
                   System.out.print(evaluator.evaluate(cell).getNumberValue());
                   break;
               case NUMERIC:
                   System.out.print(cell.getNumericCellValue());
                   System.out.print("\t");
                   break;
               case STRING:
                   System.out.print(cell.getStringCellValue());
                   System.out.print("\t");
                   break;
               case ERROR:
                   System.out.print("!");
                   System.out.print("\t");
                   break;
               }
 
           }
           System.out.println("");
       }
   }
 
}