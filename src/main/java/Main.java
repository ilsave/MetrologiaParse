import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Date;

public class Main {
    public static void main(String[] args) throws IOException {
        System.out.println("hello world");
        readFromExcel("C:\\Users\\дом\\IdeaProjects\\MetrologiaMaven\\Metrologia3.xlsx");
    }

    public static void readFromExcel(String file) throws IOException {
        XSSFWorkbook myExcelBook = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet myExcelSheet = myExcelBook.getSheet("Лист1");

        for (int i = 0; i < 44 ; i++){
            XSSFRow row = myExcelSheet.getRow(i);
            int q = i+1;
            System.out.println("Резистор №" + q + " A=|X пр−X ( n )| ");
            double middle = roundAvoid(row.getCell(0).getNumericCellValue(), 2);
            for (int j=1; j < 7; j++){
                double a = roundAvoid(Math.abs(row.getCell(j).getNumericCellValue() - middle),2);
                System.out.print("A"+j+" = |"+row.getCell(j).getNumericCellValue()+ "-" + middle + "| =  "+a);
                if (a < 928.69){
                    System.out.println(" < 928.69 - не является промахом");
                }
                if (a > 928.69){
                    System.out.println(" > 928.69 - выявлен промах");
                }
            }
           // System.out.println(row.getCell(i).getNumericCellValue());
        }

//        if(row.getCell(0).getCellType() == HSSFCell.CELL_TYPE_STRING){
//            String name = row.getCell(0).getStringCellValue();
//            System.out.println("name : " + name);
//        }
//
//        if(row.getCell(1).getCellType() == HSSFCell.CELL_TYPE_NUMERIC){
//            //Date birthdate = row.getCell(1).getDateCellValue();
//            double birthdate = row.getCell(1).getNumericCellValue();
//            System.out.println("birthdate :" + birthdate);
//        }

        myExcelBook.close();

    }
    public static double roundAvoid(double value, int places) {
        double scale = Math.pow(10, places);
        return Math.round(value * scale) / scale;
    }
}
