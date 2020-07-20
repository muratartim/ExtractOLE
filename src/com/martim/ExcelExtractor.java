package com.martim;

import com.spire.xls.OleObjectType;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import com.spire.xls.collections.WorksheetsCollection;
import com.spire.xls.core.IOleObject;

import java.nio.file.Path;

/**
 * Extracts embedded OLE objects from the source MS Excel file.
 */
public class ExcelExtractor implements Extractor {

    @Override
    public void extractOLEObjects(Path source) throws Exception {

        // create a Workbook instance
        Workbook workbook = new Workbook();

        // load the Excel document
        workbook.loadFromFile(source.toString());

        // embedded object counter
        int k = 0;

        // loop over worksheets
        WorksheetsCollection worksheets = workbook.getWorksheets();
        for (int i = 0; i < worksheets.size(); i++) {

            // get worksheet
            Worksheet sheet = worksheets.get(i);

            // has ole objects
            if (sheet.hasOleObjects()) {

                // loop over ole objects
                for (int j = 0; j < sheet.getOleObjects().size(); j++) {

                    // increment embedded object counter
                    k++;

                    // get ole object
                    IOleObject object = sheet.getOleObjects().get(j);

                    // get type of ole object
                    OleObjectType type = object.getObjectType();

                    // switch on type
                    switch (type) {
                        // word document
                        case WordDocument:
                            String doc = Extractor.getStrippedFilename(source).concat("_extracted_" + k + ".docx");
                            Extractor.byteArrayToFile(object.getOleData(), source.resolveSibling(doc).toString());
                            Main.LOGGER.info("Extracted embedded object: " + doc.toString());
                            break;
                        // powerPoint document
                        case PowerPointPresentation:
                            doc = Extractor.getStrippedFilename(source).concat("_extracted_" + k + ".pptx");
                            Extractor.byteArrayToFile(object.getOleData(), source.resolveSibling(doc).toString());
                            Main.LOGGER.info("Extracted embedded object: " + doc.toString());
                            break;
                        // pdf document
                        case Package:
                            doc = Extractor.getStrippedFilename(source).concat("_extracted_" + k + ".pdf");
                            Extractor.byteArrayToFile(object.getOleData(), source.resolveSibling(doc).toString());
                            Main.LOGGER.info("Extracted embedded object: " + doc.toString());
                            break;
                        // excel document
                        case ExcelWorksheet:
                            doc = Extractor.getStrippedFilename(source).concat("_extracted_" + k + ".xlsx");
                            Extractor.byteArrayToFile(object.getOleData(), source.resolveSibling(doc).toString());
                            Main.LOGGER.info("Extracted embedded object: " + doc.toString());
                            break;
                        default:
                            Main.LOGGER.warning("Unknown embedded object type encountered: " + object.getObjectType() + ". The object is not extracted, please extract it manually.");
                            break;
                    }
                }
            }
        }
    }
}