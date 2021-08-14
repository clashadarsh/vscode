package Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Bomb {
    public void writeExcel(String filePath,String fileName,String sheetName,String[] dataToWrite) throws IOException{
            
        File file =  new File(filePath+"\\"+fileName);
        //Create an object of FileInputStream class to read excel file
        FileInputStream inputStream = new FileInputStream(file);
    
        XSSFWorkbook workbook= new XSSFWorkbook(inputStream);
        //Read sheet inside the workbook by its name
        XSSFSheet sheet = workbook.getSheet(sheetName);
        sheet.getRow(0).createCell(2).setCellValue("bulljalcjk");
        //Create a new row and append it at last of sheet
        sheet.getRow(3).createCell(1).setCellValue("12344");        
        inputStream.close();
        //Create an object of FileOutputStream class to create write data in excel file
        FileOutputStream outputStream = new FileOutputStream(file);
        //write data in the excel file
        workbook.write(outputStream);
        //close output stream
        outputStream.close();
        workbook.close();
        }

        public static void main(String...strings) throws IOException{
            //Create an array with the data in the same order in which you expect to be filled in excel file
            String[] valueToWrite = {"Mr. E","Noida"};
            //Create an object of current class
            Bomb objExcelFile = new Bomb();
            //Write the file using file name, sheet name and the data to be filled
            objExcelFile.writeExcel("C:\\Users\\f3c3xrt\\Music","Export.xlsx","Sheet1",valueToWrite);

        }

}