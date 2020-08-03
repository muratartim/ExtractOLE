package com.martim;

import com.spire.doc.Document;
import com.spire.doc.DocumentObject;
import com.spire.doc.Section;
import com.spire.doc.collections.DocumentObjectCollection;
import com.spire.doc.collections.SectionCollection;
import com.spire.doc.documents.DocumentObjectType;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.fields.DocOleObject;

import java.nio.file.Path;

/**
 * Extracts embedded OLE objects from the source MS Word file.
 */
public class WordExtractor implements Extractor {

    @Override
    public void extractOLEObjects(Path source) throws Exception {

        // create document
        Document doc = new Document();
        doc.loadFromFile(source.toString());

        // traverse through all sections of the word document
        int count = 0;
        SectionCollection sections = doc.getSections();
        for (int i = 0; i < sections.getCount(); i++) {

            // traverse through all Child Objects in the body of each section
            Section section = sections.get(i);
            DocumentObjectCollection childObjects = section.getBody().getChildObjects();
            for (int j = 0; j < childObjects.getCount(); j++) {
                if(childObjects.get(j) instanceof Paragraph){

                    //Traverse through all Child Objects in Paragraph
                    Paragraph par = (Paragraph) childObjects.get(j);
                    for (int k = 0; k < par.getChildObjects().getCount(); k++) {

                        // find the Ole Objects and Extract
                        DocumentObject documentObject = par.getChildObjects().get(k);
                        if (documentObject.getDocumentObjectType().equals(DocumentObjectType.Ole_Object)) {

                            // increment object count
                            count++;

                            // get ole object and its type
                            DocOleObject ole = (DocOleObject) documentObject;
                            String s = ole.getObjectType();

                            // switch on type
                            switch (s) {
                                // excel document
                                case "Excel.Sheet.12":
                                    String document = Extractor.getStrippedFilename(source).concat("_extracted_" + count + ".xlsx");
                                    Extractor.byteArrayToFile(ole.getNativeData(), source.resolveSibling(document).toString());
                                    Main.LOGGER.info("Extracted embedded object: " + document);
                                    break;
                                // powerPoint document
                                case "PowerPoint.Show.12":
                                    document = Extractor.getStrippedFilename(source).concat("_extracted_" + count + ".pptx");
                                    Extractor.byteArrayToFile(ole.getNativeData(), source.resolveSibling(document).toString());
                                    Main.LOGGER.info("Extracted embedded object: " + document);
                                    break;
                                // pdf document
                                case "Package":
                                    document = Extractor.getStrippedFilename(source).concat("_extracted_" + count + ".pdf");
                                    Extractor.byteArrayToFile(ole.getNativeData(), source.resolveSibling(document).toString());
                                    Main.LOGGER.info("Extracted embedded object: " + document);
                                    break;
                                // word document
                                case "Word.Document.12":
                                    document = Extractor.getStrippedFilename(source).concat("_extracted_" + count + ".docx");
                                    Extractor.byteArrayToFile(ole.getNativeData(), source.resolveSibling(document).toString());
                                    Main.LOGGER.info("Extracted embedded object: " + document);
                                    break;
                                default:
                                    Main.LOGGER.warning("Unknown embedded object type encountered: " + s + ". The object is not extracted, please extract it manually.");
                                    break;
                            }
                        }
                    }
                }
            }
        }
    }
}
