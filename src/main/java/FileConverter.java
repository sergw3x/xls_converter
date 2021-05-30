import java.io.*;
import java.util.Arrays;

public class FileConverter {

    public static String im = "C:\\Program Files\\ImageMagick-7.0.10-Q16-HDRI\\convert.exe";

    public static boolean Convert(File fromFile, File toFile) throws IOException {

        String[] commands = {"\"" + FileConverter.im + "\"", fromFile.getAbsolutePath(), toFile.getAbsolutePath()};

        System.out.println(Arrays.toString(commands));

        return exec(commands);

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
