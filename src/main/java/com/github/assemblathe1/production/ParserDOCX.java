package com.github.assemblathe1.production;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

public class ParserDOCX {
    private final Path templateDOCXFile;
    private final List<String> keys = List.of("organisation", "number", "address");
    private final List<ObjectToInsert> allObjectsForParsing;

    public ParserDOCX(Path templateDOCXFile, Path sourseXLXSTable) throws IOException {
        this.templateDOCXFile = templateDOCXFile;
        allObjectsForParsing = getObjectsForParsing(sourseXLXSTable);
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

    public void createSpecificXWPFDocuments(Path directoryToSave) {
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
    }

    private void insertTextValueIntoXWPFRun(ObjectToInsert currentObjectToInsert, XWPFRun xwpfRun) {
        String text = xwpfRun.getText(0);
        keys.forEach(key -> {
            if (text != null) {
                if (text.contains(key)) {
                    replaceText(xwpfRun, text, key, currentObjectToInsert);
                }
            }
        });
    }

    private void replaceText(XWPFRun xwpfRun, String text, String key, ObjectToInsert currentObjectToInsert) {
        if (Objects.equals(key, "organisation")) {
            text = text.replace(key, currentObjectToInsert.getOrganisation());
        } else if (Objects.equals(key, "number")) {
            text = text.replace(key, currentObjectToInsert.getNumber().toString());
        } else if (Objects.equals(key, "address")) {
            text = text.replace(key, currentObjectToInsert.getAddress());
        }
        xwpfRun.setText(text, 0);
    }

    public List<ObjectToInsert> getObjectsForParsing(Path sourseXLXSTable) throws IOException {
        List<ObjectToInsert> allObjectsForParsing = new ArrayList<>();
        FileInputStream file = new FileInputStream(String.valueOf(sourseXLXSTable));
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell number = row.getCell(0);
            Cell organisation = row.getCell(1);
            Cell address = row.getCell(2);
            allObjectsForParsing.add(new ObjectToInsert(number.getNumericCellValue(), organisation.getStringCellValue(), address.getStringCellValue()));
        }
        return allObjectsForParsing;
    }

    private void writeDOCXFileToDisk(Path directoryToSave, XWPFDocument docxFile, String name) throws IOException {
        FileOutputStream outputStream = new FileOutputStream(directoryToSave + "\\" + name + ".docx");
        docxFile.write(outputStream);
        outputStream.close();
    }
}
