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
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author itsga
 */
public class ExcelAdapter {

    private static final String FILE_NAME = System.getProperty("user.dir") + "/books.xlsx";

    public static void updateExcelSheet(Map<String, String> map) throws IOException {
        FileOutputStream outputStream = null;
        Workbook workbook = ExcelAdapter.getExcelWorkbook();
        Sheet sheet = ExcelAdapter.getExcelSheet();
        for (Map.Entry<String, String> entry : map.entrySet()) {
            System.out.print("\tSetting " + entry.getKey() + ": " + entry.getValue() + " for ISBN = " + map.get("isbn")+"...");
            sheet.getRow(ExcelAdapter.getRowNum(map.get("isbn"))).createCell(ExcelAdapter.getColNum(entry.getKey())).setCellValue(entry.getValue());
            System.out.println("Done");
        }
        sheet.getRow(ExcelAdapter.getRowNum(map.get("isbn"))).createCell(ExcelAdapter.getColNum("sku")).setCellValue(map.get("isbn"));
        try {
            outputStream = new FileOutputStream(FILE_NAME);
            workbook.write(outputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally{
            if(workbook!=null){
                workbook.close();
            }
            if(outputStream!=null){
                outputStream.close();
            }
        }
    }

    public static List<String> getIsbnList() throws IOException {
        List<String> isbnList = new ArrayList<String>();
        Sheet sheet = getExcelSheet();
        DataFormatter formatter = new DataFormatter();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            isbnList.add(formatter.formatCellValue(sheet.getRow(i).getCell(ExcelAdapter.getColNum("isbn"))));
        }
        return isbnList;
    }

    public static Workbook getExcelWorkbook() throws FileNotFoundException, IOException {
        FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
        Workbook workbook = new XSSFWorkbook(excelFile);
        return workbook;
    }

    public static Sheet getExcelSheet() throws FileNotFoundException, IOException {
        Workbook workbook = ExcelAdapter.getExcelWorkbook();
        Sheet sheet = workbook.getSheetAt(0);
        return sheet;
    }

    public static int getRowNum(String isbn) throws IOException {
        Sheet sheet = ExcelAdapter.getExcelSheet();
        DataFormatter formatter = new DataFormatter();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            if (formatter.formatCellValue(sheet.getRow(i).getCell(getColNum("isbn"))).equalsIgnoreCase(isbn)) {
                return i;
            }
        }
        return -1;
    }

    public static int getColNum(String header) throws IOException {
        Sheet sheet = ExcelAdapter.getExcelSheet();
        DataFormatter formatter = new DataFormatter();
        for (int i = 0; i < sheet.getRow(0).getLastCellNum(); i++) {
            if (formatter.formatCellValue(sheet.getRow(0).getCell(i)).equalsIgnoreCase(header)) {
                return i;
            }
        }
        return -1;
    }
}
