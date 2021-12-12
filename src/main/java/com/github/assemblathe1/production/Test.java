package com.github.assemblathe1.production;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;

public class Test {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        Path templateDOCX = Paths.get("C:\\in\\1.docx");
        Path sourseXLXSTable = Paths.get("C:\\in\\Objects.xlsx");
        Path destinationFolder = Paths.get("C:\\in");
        new ParserDOCX(templateDOCX, sourseXLXSTable).createSpecificXWPFDocuments(destinationFolder);
    }
}
