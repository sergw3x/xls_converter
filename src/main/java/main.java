import org.apache.poi.ddf.*;
import org.apache.poi.hssf.record.EscherAggregate;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class main {

    public static void main(String[] args) throws IOException {

        String readFile = "SC9DK270G3-DBLK3448.xls";

        XLS x = new XLS();
        x.ReadFile(readFile);

//        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(readFile));
//        List lst = workbook.getAllPictures();
//        int i = 0;
//        for (Iterator it = lst.iterator(); it.hasNext(); ) {
//            PictureData pict = (PictureData)it.next();
//            String ext = pict.suggestFileExtension();
//            byte[] data = pict.getData();
////            if (ext.equals("jpeg")) {
//                FileOutputStream out = new FileOutputStream("data/tmp/pict"+i+"."+ext);
//                out.write(data);
//                out.close();
////            }
//            i++;
//        }


    }
}
