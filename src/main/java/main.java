import java.io.*;

public class main {

    public static File[] findFiles(){
        File f = new File("import");
        File[] matchingFiles = f.listFiles(new FilenameFilter() {
            public boolean accept(File dir, String name) {
                return name.endsWith("xls");
            }
        });
        return matchingFiles;
    }

    public static void main(String[] args) throws IOException {

        File[] res = findFiles();
        for (File file:res) {
//            String readFile = "import/SC9DK270G3-DBLK3448.xls";
            XLS x = new XLS(file);
            x.ReadFile();
        }

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
