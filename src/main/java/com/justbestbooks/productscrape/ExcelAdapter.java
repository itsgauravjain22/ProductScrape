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
import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import static org.apache.commons.io.FileUtils.*;
import org.apache.poi.xssf.usermodel.XSSFCell;

/**
 *
 * @author itsga
 */
public class ExcelAdapter {

    private String FILE_NAME = null;
    private final String TEST_FILE_NAME = System.getProperty("user.dir") + "/test.xlsm";
    XSSFWorkbook workbook = null;
    FileInputStream fileInputStream = null;
    FileOutputStream outputStream = null;

    public ExcelAdapter(String platform) throws IOException, InvalidFormatException {
        FILE_NAME = System.getProperty("user.dir")+"/"+platform+".xlsm";
        fileInputStream = new FileInputStream(FILE_NAME);
        workbook = new XSSFWorkbook(fileInputStream);
    }

    public void updateExcelSheet(Map<String, String> map) throws IOException {
        Sheet sheet = getExcelSheet();
        Row row;
        Cell cell;
        row = sheet.getRow(getRowNum(map.get("isbn")));
        for (Map.Entry<String, String> entry : map.entrySet()) {
            System.out.print("\tSetting " + entry.getKey() + ": " + entry.getValue().replace("\n", "") + " for ISBN = " + map.get("isbn") + "...");
            cell = row.getCell(getColNum(entry.getKey()));
            if (cell == null) {
                cell = row.createCell(getColNum(entry.getKey()));
            }
            cell.setCellValue(entry.getValue().replace("\n", ""));
            System.out.println("Done");
        }
        System.out.print("\tSetting sku: " + map.get("isbn") + " for ISBN = " + map.get("isbn") + "...");
        cell = row.getCell(getColNum("sku"));
        if (cell == null) {
            cell = row.createCell(getColNum("sku"));
        }
        cell.setCellValue(map.get("isbn"));
        System.out.println("Done");
        
        System.out.print("\tSetting published: 0 for ISBN = " + map.get("isbn") + "...");
        cell = row.getCell(getColNum("published"));
        if (cell == null) {
            cell = row.createCell(getColNum("published"));
        }
        cell.setCellValue("0");
        System.out.println("Done");
        try {
            outputStream = new FileOutputStream(new File(TEST_FILE_NAME));
            workbook.write(outputStream);
            FileUtils.copyFile(new File(TEST_FILE_NAME), new File(FILE_NAME));
            FileUtils.forceDelete(new File(TEST_FILE_NAME));
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
        Row row;
        Cell cell;
        DataFormatter formatter = new DataFormatter();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            row = sheet.getRow(i);
            cell = row.getCell(getColNum("isbn"));
            cell.setCellType(XSSFCell.CELL_TYPE_STRING);
            isbnList.add(formatter.formatCellValue(cell));
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
