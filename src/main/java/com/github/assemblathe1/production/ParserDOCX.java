package com.github.assemblathe1.production;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

public class ParserDOCX {
    private final Path templateDOCXFile;
//    private final Path numbersListTXT;
//    private final Path organisationsListTXT;
//    private final Path addressesListTXT;
    private final List<String> keys = List.of("organisation", "number", "address");
    private List<ObjectToInsert> allObjectsForParsing;

    public ParserDOCX(Path templateDOCXFile, Path numbersListTXT, Path organisationsListTXT, Path addressesListTXT) throws FileNotFoundException {
        this.templateDOCXFile = templateDOCXFile;
//        this.numbersListTXT = numbersListTXT;
//        this.organisationsListTXT = organisationsListTXT;
//        this.addressesListTXT = addressesListTXT;
        allObjectsForParsing = getObjectsForParsing(numbersListTXT, organisationsListTXT, addressesListTXT);
    }

    private XWPFDocument getTemplateDOCXFile(Path templateDOCXFile) throws IOException, InvalidFormatException {
        FileInputStream fis = new FileInputStream(String.valueOf(templateDOCXFile));
        return new XWPFDocument(OPCPackage.open(fis));
    }

    private List<String> createListInsertValues(Path valuesListTXT) throws FileNotFoundException {
        FileInputStream fileInputStream = new FileInputStream(String.valueOf(valuesListTXT));
        BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(fileInputStream));
        return bufferedReader.lines().collect(Collectors.toList());
    }

    public boolean createSpecificXWPFDocuments(Path directoryToSave) {
        //            createListInsertValues(valuesListTXT).forEach(value -> {
        allObjectsForParsing.forEach(currentObjectToInsert -> {
            try {
                XWPFDocument templateXWPFDocument = getTemplateDOCXFile(templateDOCXFile);
                templateXWPFDocument.getParagraphs().forEach(xwpfParagraph -> {
                    xwpfParagraph.getRuns().forEach(xwpfRun -> {
                                insertTextValueIntoXWPFRun(currentObjectToInsert, xwpfRun);
                            }
                    );
                });
                writeDOCXFileToDisk(directoryToSave, templateXWPFDocument, currentObjectToInsert.getOrganisation());

            } catch (IOException | InvalidFormatException e) {
                e.printStackTrace();
            }
        });
        return true;
    }

    private void insertTextValueIntoXWPFRun(ObjectToInsert currentObjectToInsert, XWPFRun xwpfRun) {
        String text = xwpfRun.getText(0);
//        System.out.println(text);
        keys.forEach(key -> {
            if (text != null) {
                if (text.contains(key)) {
                    replaceText(xwpfRun, text, key, currentObjectToInsert);
                }
            }
        });

    }

    private void replaceText(XWPFRun xwpfRun, String text, String key, ObjectToInsert currentObjectToInsert) {
        if (key == "organisation") {
            text = text.replace(key, currentObjectToInsert.getOrganisation());
        } else if (key == "number") {
            text = text.replace(key, currentObjectToInsert.getNumber());
        } else if (key == "address") {
            text = text.replace(key, currentObjectToInsert.getAddress());
        }

        xwpfRun.setText(text, 0);

    }

    public List<ObjectToInsert> getObjectsForParsing(Path numbers, Path organisations, Path addresses) throws FileNotFoundException {
        List<ObjectToInsert> allObjectsForParsing = new ArrayList<>();
        List<String> numbersList = createListInsertValues(numbers);
        List<String> organisationList = createListInsertValues(organisations);
        List<String> addressesList = createListInsertValues(addresses);
        for (int i = 0; i < numbersList.size(); i++) {
            allObjectsForParsing.add(new ObjectToInsert(numbersList.get(i), organisationList.get(i), addressesList.get(i)));
        }
        return allObjectsForParsing;
    }

    private void writeDOCXFileToDisk(Path directoryToSave, XWPFDocument docxFile, String name) throws IOException {
        FileOutputStream outputStream = new FileOutputStream(directoryToSave + "\\" + name + ".docx");
        docxFile.write(outputStream);
        outputStream.close();
    }
}
