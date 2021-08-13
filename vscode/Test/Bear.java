package Test;

import java.io.File;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.LinkedHashSet;

import org.apache.poi.xssf.usermodel.*;

public class Bear {
            public static void main(String[] args) throws IOException {

                String filePath = "C:\\Users\\f3c3xrt\\Music";
                String fileName = "Book.xlsx";
                String sheetName = "Sheet1";
                File file =    new File(filePath+"\\"+fileName);
                //Create an object of FileInputStream class to read excel file
                FileInputStream inputStream = new FileInputStream(file);
                XSSFWorkbook workbook= new XSSFWorkbook(inputStream);
                //Read sheet inside the workbook by its name
                XSSFSheet sheet = workbook.getSheet(sheetName);
                //Find number of rows in excel file
                LinkedHashSet<String> hash_Set = new LinkedHashSet<String>();
                int rowCount = sheet.getLastRowNum();
                int cellCount = sheet.getRow(1).getLastCellNum();
                //Create a loop over all the rows of excel file to read it
                for (int r = 0; r <= rowCount; r++) {
                    XSSFRow row = sheet.getRow(r);
                    for (int c = 0; c <cellCount; c++) {
                        XSSFCell cell = row.getCell(c);
                        String var = cell.getStringCellValue();
                        hash_Set.add(var); 
                    }
                } 
                System.out.println();
                System.out.println(hash_Set);
                workbook.close();
    
            }

}