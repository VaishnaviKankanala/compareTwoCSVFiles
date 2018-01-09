/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package comparisonofexcelsheets;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Scanner;
import java.util.Set;
import java.util.TreeSet;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

/**
 *
 * @author s5216
 */
public class ComparisonOfExcelSheets {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws FileNotFoundException, IOException, InvalidFormatException, JSONException {

        // TODO code application logic here
        ComparisonOfExcelSheets c = new ComparisonOfExcelSheets();

        Scanner sc = new Scanner(System.in);
        System.out.print("Excel sheet1 link: ");
        String excel1 = sc.next();
        System.out.print("Excel sheet2 link: ");
        String excel2 = sc.next();
        System.out.print("Enter column name: ");
        String colName = sc.next().toLowerCase();
        System.out.println(c.excelToJsonForSheet1(excel1, colName));
        System.out.println(c.excelToJsonForSheet2(excel2, colName));
        System.out.println(c.comparingTwoSheets(c.excelToJsonForSheet1(excel1, colName), c.excelToJsonForSheet2(excel2, colName)));

    }

    public JSONArray excelToJsonForSheet1(String excelSheet, String colName) throws FileNotFoundException, IOException, InvalidFormatException, JSONException {
        File file = new File(excelSheet);
        FileInputStream inp = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(inp);
        XSSFSheet sheet = workbook.getSheetAt(0);
        JSONObject json = new JSONObject();
        JSONArray rows = new JSONArray();

        Boolean check = false;
        int colNumber = 0;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            JSONObject jRow = new JSONObject();
            JSONArray cells = new JSONArray();
            String name = colName;
            for (int j = 0; j < sheet.getRow(i).getLastCellNum(); j++) {
                Cell cell = sheet.getRow(i).getCell(j);
                if (i == 0 && cell.getStringCellValue().toLowerCase().equals(colName)) {
                    check = true;
                    colNumber = j;
                }
                if (check == true) {
                    name = sheet.getRow(i).getCell(colNumber).getStringCellValue();
                    switch (cell.getCellTypeEnum()) {
                        case NUMERIC:
                            cells.put(cell.getNumericCellValue());
                            break;
                        case STRING:
                            cells.put(cell.getStringCellValue().toLowerCase());
                            break;
                        case BOOLEAN:
                            cells.put(cell.getBooleanCellValue());
                            break;
                    }
                }
            }
            jRow.put(name, cells);
            rows.put(jRow);
        }
        return rows;
    }

    public JSONArray excelToJsonForSheet2(String excelSheet, String colName) throws FileNotFoundException, IOException, InvalidFormatException, JSONException {
        File file = new File(excelSheet);
        FileInputStream inp = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(inp);
        XSSFSheet sheet = workbook.getSheetAt(0);
        JSONArray rows = new JSONArray();

        Boolean check = false;
        int colNumber = 0;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            JSONObject jRow = new JSONObject();
            JSONArray cells = new JSONArray();
            String name = colName;
            for (int j = 0; j < sheet.getRow(i).getLastCellNum(); j++) {
                Cell cell = sheet.getRow(i).getCell(j);
                if (i == 0 && cell.getStringCellValue().toLowerCase().equals(colName)) {
                    check = true;
                    colNumber = j;
                }
                if (check == true) {
                    name = sheet.getRow(i).getCell(colNumber).getStringCellValue();
                    switch (cell.getCellTypeEnum()) {
                        case NUMERIC:
                            cells.put(cell.getNumericCellValue());
                            break;
                        case STRING:
                            cells.put(cell.getStringCellValue().toLowerCase());
                            break;
                        case BOOLEAN:
                            cells.put(cell.getBooleanCellValue());
                            break;
                    }
                }
            }
            jRow.put(name, cells);
            rows.put(jRow);
        }
        return rows;
    }

    public JSONArray comparingTwoSheets(JSONArray sheet1data, JSONArray sheet2data) throws JSONException, FileNotFoundException {
        JSONObject json = new JSONObject();
        ArrayList<String> keysForSheet1 = new ArrayList<>();
        ArrayList<String> keysForSheet2 = new ArrayList<>();
        Set<String> resultSetKeys = new TreeSet<>();

        for (int i = 0; i < sheet1data.length(); i++) {
            json = sheet1data.getJSONObject(i);
            Iterator<String> keys = json.keys();
            while (keys.hasNext()) {
                String key = (String) keys.next();
                keysForSheet1.add(key);
            }
        }
//        System.out.println("Sheet1 size: " + keysForSheet1.size());
//        for (String s1 : keysForSheet1) {
//            System.out.println("sheet1: " + s1);
//        }

        for (int i = 0; i < sheet2data.length(); i++) {
            json = sheet2data.getJSONObject(i);
            Iterator<String> keys = json.keys();
            while (keys.hasNext()) {
                String key = (String) keys.next();
                keysForSheet2.add(key);
            }
        }
//        System.out.println("Sheet2 size: " + keysForSheet2.size());
//        for (String s2 : keysForSheet2) {
//            System.out.println("sheet1: " + s2);
//        }

        for (String keys : keysForSheet1) {
            if (!keysForSheet2.contains(keys)) {
                resultSetKeys.add(keys);
            }
        }

        for (String keys : keysForSheet2) {
            if (!keysForSheet1.contains(keys)) {
                resultSetKeys.add(keys);
            }
        }

        System.out.println("Result set: " + resultSetKeys);
        JSONObject item = new JSONObject();
        JSONArray rows = new JSONArray();
        for (String value : resultSetKeys) {
            for (int i = 0; i < sheet1data.length(); i++) {
                item = sheet1data.getJSONObject(i);
                Iterator<String> keys = item.keys();
                while (keys.hasNext()) {
                    String key = (String) keys.next();
                    if (value.equals(key)) {
                        rows.put(item.getJSONArray(value));
                    }
                }
            }
        }
        writeDataToNewExcelFile(rows);
        return rows;
    }

    private static void writeDataToNewExcelFile(JSONArray resultData) throws FileNotFoundException, JSONException {
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(
                    "C:\\jee-latest\\xslsheets\\result.csv");

            XSSFWorkbook workBook = new XSSFWorkbook();
            XSSFSheet spreadSheet = workBook.createSheet("Data");
            XSSFRow row;
            XSSFCell cell;

            int rowNum = 0;
            System.out.println("Creating excel");

            for (int i = 0; i < resultData.length(); i++) {
                row = spreadSheet.createRow(rowNum++);
                int colNum = 0;
                for (int j = 0; j < resultData.getJSONArray(i).length(); j++) {
                    cell = row.createCell(colNum++);
                    if (resultData.getJSONArray(i).get(j) instanceof String) {
                        cell.setCellValue((String) resultData.getJSONArray(i).get(j));
                    } else if (resultData.getJSONArray(i).get(j) instanceof Number) {
                        cell.setCellValue((double) (Number) resultData.getJSONArray(i).get(j));
                    }
                }
            }
            workBook.write(fos);
            // System.out.println(resultData + "--------------------");
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        } catch (IOException ex) {
            ex.printStackTrace();
        }

    }
}
