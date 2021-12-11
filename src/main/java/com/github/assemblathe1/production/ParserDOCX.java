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
    private final Path valuesListTXT;
    private final Path templateDOCXFile;
    private final List<String> keys = List.of("organisation", "number");

    public ParserDOCX(Path templateDOCXFile, Path valuesListTXT) throws IOException, InvalidFormatException {
        this.valuesListTXT = valuesListTXT;
        this.templateDOCXFile = templateDOCXFile;

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
        try {
            createListInsertValues(valuesListTXT).forEach(value -> {
                try {
                    XWPFDocument templateXWPFDocument = getTemplateDOCXFile(templateDOCXFile);
                    templateXWPFDocument.getParagraphs().forEach(xwpfParagraph -> {
                        xwpfParagraph.getRuns().forEach(xwpfRun -> {
                                     insertTextValueIntoXWPFRun(value, xwpfRun);
                                }
                        );
                    });
                    writeDOCXFileToDisk(directoryToSave, templateXWPFDocument, value);

                } catch (IOException | InvalidFormatException e) {
                    e.printStackTrace();
                }
            });
            return true;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            return false;
        }
    }

    private void insertTextValueIntoXWPFRun(String value, XWPFRun xwpfRun) {
        String text = xwpfRun.getText(0);
//        System.out.println(text);
        keys.forEach(key -> {
            if (text != null) {
                if (text.contains(key)) {
                    replaceText(value, xwpfRun, text, key);
                }
            }
        });

    }

    private void replaceText(String value, XWPFRun xwpfRun, String key, String text) {
        text = text.replace(key, value);
        xwpfRun.setText(text, 0);
    }

    private void writeDOCXFileToDisk(Path directoryToSave, XWPFDocument docxFile, String name) throws IOException {
        FileOutputStream outputStream = new FileOutputStream(directoryToSave + "\\" + name + ".docx");
        docxFile.write(outputStream);
        outputStream.close();
    }
}
