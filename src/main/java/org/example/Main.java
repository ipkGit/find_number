package org.example;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

public class Main {
    public static void main(String[] args) {
        String result = CarNumberService.findCarNumberInXLSX("name_java.xlsx", "0Мо");
        System.out.println(result);
    }
}

class CarNumberService {
    public static String findCarNumberInXLSX(String file, String toFind) {

        try (InputStream in = new FileInputStream(file)) {
            Workbook workbook = WorkbookFactory.create(in);
            if (findNumber(workbook, toFind.toUpperCase())) {
                return "Номер найден";
            }
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
        return "Номер не найден";
    }

    private static boolean findNumber(Workbook workbook, String toFind) {
        Sheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(0);
            if (cell.getCellType() == CellType.STRING) {
                String value = cell.getStringCellValue();
                if (value.length() > 7) {
                    value = cell.getStringCellValue().substring(0, 6);
                    if (value.contains(toFind)) {
                        return true;
                    }
                }
            }
        }
        return false;
    }
}

