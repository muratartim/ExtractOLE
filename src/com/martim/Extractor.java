package com.martim;

import java.io.*;
import java.nio.file.Path;
import java.util.logging.FileHandler;
import java.util.logging.Level;
import java.util.logging.Logger;

public interface Extractor {

    /**
     * Returns stripped file name of the given path.
     *
     * @param source Path to source file.
     * @return The stripped file name.
     */
    static String getStrippedFilename(Path source) {
        String sourceFilename = source.getFileName().toString();
        return sourceFilename.substring(0, sourceFilename.lastIndexOf("."));
    }

    /**
     * Writes the given byte array to file.
     *
     * @param data     Byte array to write.
     * @param destPath Path to destination file.
     */
    static void byteArrayToFile(byte[] data, String destPath) {
        File dest = new File(destPath);
        try (InputStream is = new ByteArrayInputStream(data);
             OutputStream os = new BufferedOutputStream(new FileOutputStream(dest, false));) {
            byte[] flush = new byte[1024];
            int len = -1;
            while ((len = is.read(flush)) != -1) {
                os.write(flush, 0, len);
            }
            os.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Extracts OLE objects from the given file. Extracted files are written as siblings of the input file with a standard naming convention.
     *
     * @param source Path to input file.
     */
    void extractOLEObjects(Path source) throws Exception;

}
