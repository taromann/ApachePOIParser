package com.github.assemblathe1.preview;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadPoolExecutor;

public class ApachPOITestWriting {

    public static void main(String[] args) throws IOException {

        ExecutorService executorService = Executors.newFixedThreadPool(10);
        for (int i = 0; i < 1; i++) {
            int finalI = i;
            executorService.execute(new Runnable() {
                @Override
                public void run() {
                    try {
                        createDOCX(finalI);
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            });
        }

        executorService.shutdown();
    }

    private static void createDOCX(int i) throws IOException {
        XWPFDocument docxModel = new XWPFDocument();
        CTSectPr ctSectPr = docxModel.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(docxModel, ctSectPr);

        CTP ctpHeaderModel = createHeaderModel(
                i + " = i"
        );

        XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeaderModel, docxModel);

        headerFooterPolicy.createHeader(
                XWPFHeaderFooterPolicy.DEFAULT,
                new XWPFParagraph[]{headerParagraph});

        CTP ctpFooterModel = createFooterModel(i + " = i");

//        XWPFParagraph footerParagraph = new XWPFParagraph(ctpFooterModel, docxModel);
//        headerFooterPolicy.createFooter(
//                XWPFHeaderFooterPolicy.DEFAULT,
//                new XWPFParagraph[]{footerParagraph}
//        );

        XWPFParagraph bodyParagraph = docxModel.createParagraph();
        bodyParagraph.setAlignment(ParagraphAlignment.RIGHT);
        XWPFRun paragraphConfig = bodyParagraph.createRun();
        paragraphConfig.setItalic(true);
        paragraphConfig.setFontSize(25);
        paragraphConfig.setColor("06357a");
        paragraphConfig.setText(
                "paragraphConfig"
        );
        paragraphConfig.setFontFamily("Times New Roman");

        XWPFRun paragraphConfig_1 = bodyParagraph.createRun();
        paragraphConfig_1.setItalic(false);
        paragraphConfig_1.setFontSize(10);
        paragraphConfig_1.setText("paragraphConfig_1");

        XWPFTable xwpfTable = docxModel.createTable(4, 5);

        CTTblWidth widthRepr = xwpfTable.getCTTbl().getTblPr().addNewTblW();
        widthRepr.setType(STTblWidth.DXA);
        widthRepr.setW(BigInteger.valueOf(10000));

        xwpfTable.getRow(2).getCell(3).setParagraph(bodyParagraph);




        XWPFParagraph bodyParagraph2 = docxModel.createParagraph();
        bodyParagraph2.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun paragraphConfig2 = bodyParagraph2.createRun();
//        paragraphConfig2.setItalic(true);
        paragraphConfig2.setFontSize(12);
//        paragraphConfig2.setColor("06357a");
        paragraphConfig2.setText(
                "Paragraph 2"
        );
        paragraphConfig2.setFontFamily("Arial");


        FileOutputStream outputStream = new FileOutputStream("C:\\in\\" + i + "__" + Thread.currentThread().getName() + ".docx");
        docxModel.write(outputStream);
        outputStream.close();
        System.out.println("Успешно записан в файл");
    }

    private static CTP createFooterModel(String footerContent) {
        // создаем футер или нижний колонтитул
        CTP ctpFooterModel = CTP.Factory.newInstance();
        CTR ctrFooterModel = ctpFooterModel.addNewR();
        CTText cttFooter = ctrFooterModel.addNewT();

        cttFooter.setStringValue(footerContent);
        return ctpFooterModel;
    }

    private static CTP createHeaderModel(String headerContent) {
        // создаем хедер или верхний колонтитул
        CTP ctpHeaderModel = CTP.Factory.newInstance();
        CTR ctrHeaderModel = ctpHeaderModel.addNewR();
        CTText cttHeader = ctrHeaderModel.addNewT();

        cttHeader.setStringValue(headerContent);
        return ctpHeaderModel;
    }




}
