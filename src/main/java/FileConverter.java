import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Arrays;

public class FileConverter {

    public static String im = "C:\\Program Files\\ImageMagick-7.0.10-Q16-HDRI\\convert.exe";

    public static boolean Convert(File fromFile, File toFile) throws IOException {

        String[] commands;
        if (OSDetector.isWindows()){
            commands = new String[]{"\"" + FileConverter.im + "\"", fromFile.getAbsolutePath(), toFile.getAbsolutePath()};
        }else{
            commands = new String[]{"emf2svg-conv", "-i", fromFile.getAbsolutePath(), "-o", toFile.getAbsolutePath()};
        }
        return exec(commands);

    }

    private static Boolean CopyFile(Path from, Path to){
        try {
            Files.copy(from, to, StandardCopyOption.REPLACE_EXISTING);
            return true;
        } catch (IOException ex) {
            System.err.format("%s", ex);
            return false;
        }
    }

    private static Boolean exec(String[] commands) throws IOException {
        Runtime rt = Runtime.getRuntime();
        Process proc = rt.exec(commands);
        String s = null;

        BufferedReader stdError = new BufferedReader(new
                InputStreamReader(proc.getErrorStream()));

        if ((s = stdError.readLine()) != null) {
            System.out.println(s);
            return false;
        }
        return true;
    }
}
