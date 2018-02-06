/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.justbestbooks.productscrape;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author itsga
 */
public class ExcelAdapter {

    private final String FILE_NAME = System.getProperty("user.dir") + "/books.xlsx";
    XSSFWorkbook workbook = null;

    public ExcelAdapter() throws IOException, InvalidFormatException {
        workbook = new XSSFWorkbook(new File(FILE_NAME));
    }

    public void updateExcelSheet(Map<String, String> map) throws IOException {
        FileOutputStream outputStream = null;
        Sheet sheet = getExcelSheet();
        Row row;
        Cell cell;
        for (Map.Entry<String, String> entry : map.entrySet()) {
            System.out.print("\tSetting " + entry.getKey() + ": " + entry.getValue() + " for ISBN = " + map.get("isbn") + "...");
            row = sheet.getRow(getRowNum(map.get("isbn")));
            cell = row.getCell(getColNum(entry.getKey()));
            if (cell == null)
                cell = row.createCell(getColNum(entry.getKey()));
            cell.setCellValue(entry.getValue());
            System.out.print("\n"+row+"\n");
            System.out.println("Done");
        }
        sheet.getRow(getRowNum(map.get("isbn"))).createCell(getColNum("sku")).setCellValue(map.get("isbn"));
        try {
            outputStream = new FileOutputStream(new File(FILE_NAME));
            workbook.write(outputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (outputStream != null) {
                outputStream.close();
            }
        }
    }

    public List<String> getIsbnList() throws IOException {
        List<String> isbnList = new ArrayList<String>();
        Sheet sheet = getExcelSheet();
        DataFormatter formatter = new DataFormatter();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            isbnList.add(formatter.formatCellValue(sheet.getRow(i).getCell(getColNum("isbn"))));
        }
        return isbnList;
    }
    
    public Sheet getExcelSheet() {
        Sheet sheet = workbook.getSheetAt(0);
        return sheet;
    }

    public int getRowNum(String isbn) throws IOException {
        Sheet sheet = getExcelSheet();
        DataFormatter formatter = new DataFormatter();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            if (formatter.formatCellValue(sheet.getRow(i).getCell(getColNum("isbn"))).equalsIgnoreCase(isbn)) {
                return i;
            }
        }
        return -1;
    }

    public int getColNum(String header) throws IOException {
        Sheet sheet = getExcelSheet();
        DataFormatter formatter = new DataFormatter();
        for (int i = 0; i < sheet.getRow(0).getLastCellNum(); i++) {
            if (formatter.formatCellValue(sheet.getRow(0).getCell(i)).equalsIgnoreCase(header)) {
                return i;
            }
        }
        return -1;
    }
}
