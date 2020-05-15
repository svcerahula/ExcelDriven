package FileUtilities;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;

public class testUtilities {

    public static FileInputStream getStreamObjectForFile(String className, String fileName)
            throws FileNotFoundException, ClassNotFoundException {
        Class myClass = Class.forName(className);
        File file = new File(myClass.getResource(fileName).getFile());
        String path = file.getAbsolutePath();
        path = URLDecoder.decode(path, StandardCharsets.UTF_8); // remove encoders added to the file path
        FileInputStream fis =new FileInputStream(path);
        return fis;
    }
}
