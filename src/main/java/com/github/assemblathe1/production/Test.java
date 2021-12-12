package com.github.assemblathe1.production;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;

public class Test {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        Path templateDOCX = Paths.get("C:\\in\\1.docx");
        Path numbersListTXT = Paths.get("C:\\in\\Numbers.txt");
        Path organisationsListTXT = Paths.get("C:\\in\\Organisations.txt");
        Path addressesListTXT = Paths.get("C:\\in\\Addresses.txt");
        new ParserDOCX(templateDOCX, numbersListTXT, organisationsListTXT, addressesListTXT).createSpecificXWPFDocuments(Path.of("C:\\in"));
    }
}
