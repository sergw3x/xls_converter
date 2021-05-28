import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

public class XLS {

    @SuppressWarnings("deprecation")
    public static void writeIntoExcel(String file) throws FileNotFoundException, IOException {
        Workbook book = new HSSFWorkbook();
        Sheet sheet = book.createSheet("Birthdays");

        // Нумерация начинается с нуля
        Row row = sheet.createRow(0);

        // Мы запишем имя и дату в два столбца
        // имя будет String, а дата рождения --- Date,
        // формата dd.mm.yyyy
        Cell name = row.createCell(0);
        name.setCellValue("John");

        Cell birthdate = row.createCell(1);

        DataFormat format = book.createDataFormat();
        CellStyle dateStyle = book.createCellStyle();
        dateStyle.setDataFormat(format.getFormat("dd.mm.yyyy"));
        birthdate.setCellStyle(dateStyle);


        // Нумерация лет начинается с 1900-го
        birthdate.setCellValue(new Date(110, Calendar.OCTOBER, 10));

        // Меняем размер столбца
        sheet.autoSizeColumn(1);

        // Записываем всё в файл
        book.write(new FileOutputStream(file));
        book.close();
    }

    public static void ReadFile(String file) throws IOException {
        HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet myExcelSheet = myExcelBook.getSheet("D111A");

        int rowTotal = myExcelSheet.getLastRowNum();
        if ((rowTotal > 0) || (myExcelSheet.getPhysicalNumberOfRows() > 0)) {
            for (int rowNum = 1; rowNum < rowTotal; rowNum++){
//                System.out.println(rowTotal);
                HSSFRow row = myExcelSheet.getRow(rowNum);
                if (row == null){
                    continue;
                }
                int cellTotal = row.getLastCellNum();
                for (int cellNum = 1; cellNum < cellTotal; cellNum++){

                    if(row.getCell(cellNum).getCellType() == CellType.STRING){
                        String val = row.getCell(cellNum).getStringCellValue();
                        System.out.printf("R%sC%s: %s", rowNum, cellNum, val);
                    }

                    if(row.getCell(cellNum).getCellType() == CellType.NUMERIC){
                        Date val = row.getCell(cellNum).getDateCellValue();
                        System.out.printf("R%sC%s: %s", rowNum, cellNum, val);
                    }
                }
                System.out.print("\n");
            }
        }



        myExcelBook.close();
    }

}
