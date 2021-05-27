import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.*;
import java.util.Calendar;
import java.util.Date;

public class main {

    public static void main(String[] args) throws IOException {
        System.out.println("sss");

        String filename = "testfile.xls";
        XLS.writeIntoExcel(filename);

        OSFileRunner.open(new File(filename));
    }



}
