import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class XLS {

    private String filename;

    public void ReadFile(String file) throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(file));
        filename = XLS.getFileNameWithoutExtension(file);
        System.out.println(filename);

        HSSFSheet sheet = wb.getSheet("D111A");
        this.readTable(sheet);

        for (int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex++) {
//            HSSFSheet sheet = wb.getSheetAt(sheetIndex);
//            this.readTable(sheet); // "D111A"

//            this.saveImageFromSheet(sheet);
        }

        wb.close();
    }

    private void readTable(HSSFSheet sheet) {
        int rowTotal = sheet.getLastRowNum();
        if ((rowTotal > 0) || (sheet.getPhysicalNumberOfRows() > 0)) {
            for (int rowNum = 1; rowNum < rowTotal; rowNum++) {
                HSSFRow row = sheet.getRow(rowNum);
                if (row == null) {
                    continue;
                }
                int cellTotal = row.getLastCellNum();
                for (int cellNum = 1; cellNum < cellTotal; cellNum++) {

                    if (row.getCell(cellNum).getCellTypeEnum() == CellType.STRING) {
                        String val = row.getCell(cellNum).getStringCellValue();
                        System.out.printf("R%sC%s: %s", rowNum, cellNum, val);
                    }

                    if (row.getCell(cellNum).getCellTypeEnum() == CellType.NUMERIC) {
                        Date val = row.getCell(cellNum).getDateCellValue();
                        System.out.printf("R%sC%s: %s", rowNum, cellNum, val);
                    }
                }
                System.out.print("\n");
            }
        }

    }

    private void saveImageFromSheet(Sheet sheet) throws IOException {

        Drawing<?> draw = sheet.createDrawingPatriarch();
        List<Picture> pics = new ArrayList<>();
        if (draw instanceof HSSFPatriarch) {
            HSSFPatriarch hp = (HSSFPatriarch) draw;
            for (HSSFShape hs : hp.getChildren()) {
                if (hs instanceof Picture) {
                    pics.add((Picture) hs);
                }
            }
        } else {
            XSSFDrawing xdraw = (XSSFDrawing) draw;
            for (XSSFShape xs : xdraw.getShapes()) {
                if (xs instanceof Picture) {
                    pics.add((Picture) xs);
                }
            }
        }

        for (Picture p : pics) {
            PictureData pd = p.getPictureData();
            String ext = pd.suggestFileExtension();
            // todo: https://github.com/kakwa/libemf2svg
            this.saveFile(sheet.getSheetName(), ext, pd.getData());
        }
    }

    //Write the Excel file
    private void saveFile(String name, String ext, byte[] picData) throws IOException {

        Path tmpDir = Paths.get("data", "tmp");
        Path tmpFile = Paths.get(tmpDir.toString(), name + "." + ext );

        Path targetDir = Paths.get("data", this.filename, "img");
        Path targetFile = Paths.get(targetDir.toString(), name + ".svg");

        tmpDir.toFile().mkdirs();
        targetDir.toFile().mkdirs();

        // Сохраним во временную
        try (FileOutputStream fos = new FileOutputStream(tmpFile.toString())) {
            fos.write(picData);
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Конвертируем
        if (FileConverter.Convert(tmpFile.toFile(), targetFile.toFile()) && !tmpFile.toFile().delete()){
            System.out.println("err delete tmp file: "+tmpFile);
        }
    }

    private static String getFileNameWithoutExtension(String f) {
        if (f.contains(".")) {
            return f.replace(f.substring(f.lastIndexOf(".")), "");
        }
        return f;
    }
}
