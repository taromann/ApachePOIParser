package com.github.assemblathe1.preview;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.*;
import java.util.List;
import java.util.stream.Collectors;

public class ParsingTestDOCX {
    public static int currenObject = 0;

    public static void main(String[] args) throws IOException, InvalidFormatException {
        FileInputStream fileInputStream = new FileInputStream("C:\\in\\1.txt");
        BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(fileInputStream));
        List<String> riasObjects = bufferedReader.lines().collect(Collectors.toList());

        FileInputStream fis = new FileInputStream("C:\\in\\1.docx");
        XWPFDocument docxFile = new XWPFDocument(OPCPackage.open(fis));
        List<XWPFParagraph> paragraphs = docxFile.getParagraphs();
        riasObjects.forEach(riasObject -> {
            try {
                creatreDOCX(docxFile, paragraphs.get(19), riasObjects, currenObject);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });
    }

    public static void creatreDOCX(XWPFDocument docxFile, XWPFParagraph xwpfParagraph,  List<String> riasObjects, int currenObjectNumber) throws IOException {
        String currentObject = riasObjects.get(currenObjectNumber);
        xwpfParagraph.removeRun(0);
        xwpfParagraph.createRun().setText(currentObject);
        FileOutputStream outputStream = new FileOutputStream("C:\\in\\" + currentObject + ".docx");
        docxFile.write(outputStream);
        outputStream.close();
        currenObject++;
    }
}
