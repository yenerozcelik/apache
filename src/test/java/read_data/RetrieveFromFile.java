package read_data;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class RetrieveFromFile {

    @Test
    public void readFileTest() throws IOException {

        File excelfile=new File("src/test/resources/TestSetup.xlsx");

        FileInputStream fileInputStream=new FileInputStream(excelfile);

        XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream);

        XSSFSheet page1= workbook.getSheet("Sheet1");

        XSSFRow row1 =page1.getRow(0);

        XSSFCell cell1 = row1.getCell(0);

        System.out.println(cell1);

        XSSFRow row2 = page1.getRow(1);

        XSSFCell cell2 = row2.getCell(0);

        System.out.println(cell2);
    }

    @Test
    public void getRowValuesTest() throws IOException {

        File excelfile=new File("src/test/resources/TestSetup.xlsx");

        FileInputStream fileInputStream=new FileInputStream(excelfile);

        XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream);

        XSSFSheet sheet1= workbook.getSheetAt(0);

        XSSFRow row1 = sheet1.getRow(0);

        for (int i = row1.getFirstCellNum() ; i <row1.getLastCellNum() ; i++) {
            XSSFCell cell = row1.getCell(i);
            System.out.print(cell + "|"); //Environment|endpoint|tag|suite|status|date|

        }

    }

    @Test
    public void getAllDataTest() throws IOException {
        File excelfile = new File("src/test/resources/TestSetup.xlsx");

        FileInputStream fileInputStream = new FileInputStream(excelfile);

        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

        XSSFSheet sheet1 = workbook.getSheetAt(0);


        for (int i = sheet1.getFirstRowNum(); i < sheet1.getLastRowNum(); i++) {
            XSSFRow rows = sheet1.getRow(i);
            System.out.print("| ");
            for (int j = rows.getFirstCellNum(); j < rows.getLastCellNum(); j++) {
                XSSFCell cell = rows.getCell(j);
                System.out.print(cell + " | ");
            }
            System.out.println();
        }
    }
    }
