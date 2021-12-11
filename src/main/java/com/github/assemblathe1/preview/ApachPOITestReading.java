package com.github.assemblathe1.preview;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

public class ApachPOITestReading {
    public ApachPOITestReading() throws FileNotFoundException {
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        FileInputStream fileInputStream = new FileInputStream("C:\\in\\0__pool-1-thread-1.docx");

        // открываем файл и считываем его содержимое в объект XWPFDocument
        XWPFDocument docxFile = new XWPFDocument(OPCPackage.open(fileInputStream));
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(docxFile);

        // считываем верхний колонтитул (херед документа)
        XWPFHeader docHeader = headerFooterPolicy.getDefaultHeader();
        System.out.println(docHeader.getText());

        // печатаем содержимое всех параграфов документа в консоль
        List<XWPFParagraph> paragraphs = docxFile.getParagraphs();
        paragraphs.forEach(p -> insertText(p, p.getText().toUpperCase()));


        // считываем нижний колонтитул (футер документа)
//        XWPFFooter docFooter = headerFooterPolicy.getDefaultFooter();
//        System.out.println(docFooter.getText());

        FileOutputStream outputStream = new FileOutputStream("C:\\in\\1.docx");
        docxFile.write(outputStream);
        outputStream.close();

        System.out.println("_____________________________________");
        // печатаем все содержимое Word файла
//            XWPFWordExtractor extractor = new XWPFWordExtractor(docxFile);
//            System.out.println(extractor.getText());

    }

    public static void insertText(XWPFParagraph xwpfParagraph, String string) {
        xwpfParagraph.getRuns().stream().forEach(xwpfRun -> xwpfRun.setText(string));
    }

}
