package com.martim;

import javax.activation.UnsupportedDataTypeException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.logging.FileHandler;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

public class Main {

    public static Logger LOGGER;

    public static void main(String[] args) {

        // invalid inputs or help requested
        if (args == null || args.length == 0 || args[0] == null || args[0].equals("-help")) {
            printHelp();
            System.exit(0);
        }

        // invalid source path
        Path sourcePath = Paths.get(args[0]);
        if (!Files.isRegularFile(sourcePath)) {
            printHelp();
            System.exit(0);
        }

        // create logger
        LOGGER = createLogger(sourcePath);
        LOGGER.info("Processing input file: " + sourcePath.getFileName().toString());

        // extract ole objects
        try {
            Extractor extractor = createExtractor(sourcePath);
            extractor.extractOLEObjects(sourcePath);
        } catch (Exception e) {
            e.printStackTrace();
            LOGGER.log(Level.WARNING, "Exception occurred during extraction:", e);
        }
    }

    private static void printHelp() {
        System.out.println("Application usage:");
        System.out.println("Extract OLE objects (with sibling source file): java -jar extractOle.jar <source filename>");
        System.out.println("Extract OLE objects (with full source path): java -jar extractOle.jar <source path>");
        System.out.println("Supported source file types: *.docx, *.xlsx");
        System.out.println("Help (which prints this): java -jar extractOle.jar -help");
    }

    private static Logger createLogger(Path source) {

        try {

            // create logger
            Logger logger = Logger.getLogger(Extractor.class.getName());

            // create file handler
            FileHandler fileHandler = new FileHandler(source.resolveSibling(Extractor.getStrippedFilename(source).concat("_extracted.log")).toString());

            // set simple formatter to file handler
            fileHandler.setFormatter(new SimpleFormatter());

            // add handler to logger
            logger.addHandler(fileHandler);

            // set log level (info, warning and severe are logged)
            logger.setLevel(Level.ALL);

            // return logger
            return logger;
        }

        // exception occurred during creating logger
        catch (SecurityException | IOException e) {
            e.printStackTrace();
            System.exit(1);
            return null;
        }
    }

    /**
     * Returns the file type of the given file name, or null if file type is not recognized.
     *
     * @param fileName
     *            File name.
     * @return File type of the given file name, or null if file type is not recognized.
     */
    private static String getFileExtension(String fileName) {

        // no file extension found
        int index = fileName.lastIndexOf(".");
        if (index <= 0)
            return null;

        // get file extension
        return fileName.substring(index);
    }

    /**
     * Creates and returns the embedded OLE object extractor for the provided source input file.
     *
     * @param source Path to input file.
     * @return Newly created extractor.
     * @throws UnsupportedDataTypeException If unsupported input data type is passed.
     */
    private static Extractor createExtractor(Path source) throws UnsupportedDataTypeException {
        String ext = getFileExtension(source.getFileName().toString());
        if (ext == null) {
            throw new UnsupportedDataTypeException("Unsupported input file type: null. Please provide .docx or .xlsx input files.");
        }
        switch (ext) {
            case ".docx":
                return new WordExtractor();
            case ".xlsx":
                return new ExcelExtractor();
            default:
                throw new UnsupportedDataTypeException("Unsupported input file type: " + ext + ". Please provide .docx or .xlsx input files.");
        }
    }
}
