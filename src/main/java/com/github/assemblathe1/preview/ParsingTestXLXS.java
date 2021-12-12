package com.github.assemblathe1.preview;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ParsingTestXLXS {
    public static void main(String[] args) throws IOException {
        FileInputStream file = new FileInputStream(new File("C:\\in\\Objects.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();


        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell number = row.getCell(0);
            Cell organisation = row.getCell(1);
            Cell address = row.getCell(2);
            System.out.printf("number = %s, organisation = %s, address = %s", number.getNumericCellValue(), organisation.getStringCellValue(), address.getStringCellValue());
            System.out.println("\n");



//            Iterator<Cell> cells = row.iterator();
//            while (cells.hasNext()) {
//                Cell cell = cells.next();
//                CellType cellType = cell.getCellType();
//                switch (cellType) {
//                    case STRING:
//                        System.out.println(cell.getStringCellValue());
//                        break;
//                    case NUMERIC:
//                        System.out.println(cell.getNumericCellValue());
//                        break;
//
//                }


//            }

        }

    }
}
