import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;

public class XLS {

    private final String fileName;
    private String fileNameWithoutExt = "";

    public XLS(String file){
        this.fileName = file;

        if (file.contains(".")) {
            this.fileNameWithoutExt = file.replace(file.substring(file.lastIndexOf(".")), "");
        }
    }

    public void ReadFile() throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(this.fileName));

        System.out.println("Reading file: " + this.fileName);

        Document Obj = new Document();
        Obj.Name = "";
        Obj.Description = "";

        for (int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex++) {
            HSSFSheet sheet = wb.getSheetAt(sheetIndex);
            System.out.println("Reading sheet: " + sheet.getSheetName());
            if (sheetIndex == 0) {
                this.readSheet(sheet, Obj, true);
            } else {
                this.readSheet(sheet, Obj, false);
                this.saveImageFromSheet(sheet);
            }
        }

        wb.close();
    }

    private void readSheet(HSSFSheet sheet, Document Obj, Boolean firstSheet) {

        String prevCodeRange = "";
        String prevCode = "";

        if (Obj.Sheets == null){
            Obj.Sheets = new ArrayList<>();
        }
        ArrayList<Map<String, String>> arrayList = new ArrayList<>();
        Obj.mapColNames = new HashMap<>();

        if ((sheet.getLastRowNum() > 0) || (sheet.getPhysicalNumberOfRows() > 0)) {
            //row loop
            Iterator<Row> rowIterator = sheet.rowIterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                Iterator<Cell> cellIterator = row.cellIterator();

                Map<String, String> tabRow = new HashMap<>();

                outer:
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    String cellValue = "";
                    int cellIndex = cell.getColumnIndex();

                    if (cell.getCellType().equals(CellType.NUMERIC)) {
                        int num = (int) Math.round(cell.getNumericCellValue());
                        cellValue = Integer.toString(num);
                    } else {
                        cellValue = cell.getStringCellValue();
                    }

                    cellValue = cellValue.trim();

                    String putMapValue = "";
                    if (firstSheet) {
                        if (Obj.Name.equals("")) {
                            Obj.Name = cellValue;
                            continue;
                        }

                        if (Obj.Description.equals("")) {
                            Obj.Description = cellValue;
                            continue;
                        }

                        if (cellValue.contains("Group Number")) {
                            putMapValue = "GroupNumber";
                        } else if (cellValue.contains("Chinese Description")) {
                            putMapValue = "ChineseDescription";
                        } else if (cellValue.contains("English Description")) {
                            putMapValue = "EnglishDescription";
                        } else if (cellValue.contains("Code")) {
                            putMapValue = "Code";
                        }
                    } else {
                        if (cellValue.contains("Ref")) {
                            putMapValue = "Ref";
                        } else if (cellValue.contains("Part No")) {
                            putMapValue = "PartNo";
                        } else if (cellValue.contains("Chinese Description")) {
                            putMapValue = "ChineseDescription";
                        } else if (cellValue.contains("English Description")) {
                            putMapValue = "EnglishDescription";
                        } else if (cellValue.contains("Quantity")) {
                            putMapValue = "Quantity";
                        } else if (cellValue.contains("Standard Fasteners Sign")) {
                            putMapValue = "StandardFastenersSign";
                        }
                    }
                    if (!putMapValue.equals("")) {
                        Obj.mapColNames.put(cellIndex, putMapValue);
                        continue;
                    }

                    //will iterate over the Merged cells
                    for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                        CellRangeAddress mergedRegion = sheet.getMergedRegion(i);

                        int firstColumn = mergedRegion.getFirstColumn();
                        int firstRow = mergedRegion.getFirstRow();

                        if (firstRow == cell.getRowIndex() && firstColumn == cell.getColumnIndex()) {
                            if (prevCodeRange.equals("") || !prevCodeRange.equals(mergedRegion.formatAsString())) {
                                prevCodeRange = mergedRegion.formatAsString();
                                prevCode = cellValue;
                                tabRow.put(Obj.mapColNames.get(firstColumn), cellValue);
                            } else {
                                tabRow.put(Obj.mapColNames.get(firstColumn), prevCode);
                            }
                            continue outer;
                        }
                    }

                    //the data in merge cells is always present on the first cell.
                    // All other cells(in merged region) are considered blank
                    if (cell.getCellType() == CellType.BLANK) {
                        if (!prevCodeRange.equals("") && !Obj.mapColNames.get(cellIndex).equals("")) {
                            String[] range = prevCodeRange.split(":");
                            String min = range[0]; // G5
                            String max = range[1]; // G9
                            String colNameString = CellReference.convertNumToColString(cellIndex);

                            // G == G && 9 == 9
                            if (getColNameFromCellName(min).equals(colNameString) &&
                                    getColNameFromCellName(max).equals(colNameString) &&
                                    row.getRowNum() < getRowNumFromCellName(max) &&
                                    getRowNumFromCellName(min) <= row.getRowNum()
                            ) {
                                tabRow.put(Obj.mapColNames.get(cellIndex), prevCode);
                            }
                        }
                        continue;
                    }
                    if (cellValue.contains("Back to") && cell.getHyperlink() != null) {
                        continue;
                    }else {
                        if (Obj.mapColNames.get(cellIndex) != null && !cellValue.equals("")){
                            tabRow.put(Obj.mapColNames.get(cellIndex), cellValue);
                        }
                    }
                }
                if (!tabRow.isEmpty()) {
                    arrayList.add(tabRow);
                }
            }
        }
        Obj.Sheets.add(arrayList);

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
        Path tmpFile = Paths.get(tmpDir.toString(), name + "." + ext);

        Path targetDir = Paths.get("data", this.fileNameWithoutExt, "img");
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
        if (FileConverter.Convert(tmpFile.toFile(), targetFile.toFile()) && !tmpFile.toFile().delete()) {
            System.out.println("err delete tmp file: " + tmpFile);
        }
    }

    private static String getColNameFromCellName(String s) {
        return s.replaceAll("[^A-Za-z]", "");
        //CellReference.convertColStringToIndex(colName);
    }

    private static int getRowNumFromCellName(String s) {
        return Integer.parseInt(s.replaceAll("[^0-9]", ""));
    }

}
