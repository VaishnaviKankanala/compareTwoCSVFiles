/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package comparisonofexcelsheets;

import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;
import java.util.Set;
import java.util.TreeSet;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
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
        JSONArray file1 = c.excelToJsonForSheet(excel1, colName);
        JSONArray file2 = c.excelToJsonForSheet(excel2, colName);
        System.out.println(c.comparingTwoSheets(file1, file2, colName));

    }

    public JSONArray excelToJsonForSheet(String excelSheet, String colName) throws FileNotFoundException, IOException, InvalidFormatException, JSONException {
        CSVReader reader = new CSVReader(new FileReader(excelSheet));
        List<String[]> records = reader.readAll();
        JSONArray rows = new JSONArray();
        JSONObject jRow = new JSONObject();
        int columnNumber = 0;
        String[] a = records.get(0);
        for (int k = 0; k < a.length; k++) {
            if (colName.equals(a[k])) {
                columnNumber = k;
                System.out.println("k: " + k);
            }
        }
        String name = "";
        for (int i = 0; i < records.size(); i++) {
            String[] s = records.get(i);
            JSONArray record = new JSONArray(s);
            name = s[columnNumber];
            jRow.put(name, record);
        }
        rows.put(jRow);
        reader.close();
        return rows;
    }

    public JSONArray comparingTwoSheets(JSONArray sheet1data, JSONArray sheet2data, String colName) throws JSONException, FileNotFoundException {
        JSONObject json = new JSONObject();
        ArrayList<String> keysForSheet1 = new ArrayList<>();
        ArrayList<String> keysForSheet2 = new ArrayList<>();
        Set<String> resultSetKeys = new TreeSet<>();

        for (int i = 0; i < sheet1data.length(); i++) {
            json = sheet1data.getJSONObject(i);
            Iterator<String> keys = json.keys();
            while (keys.hasNext()) {
                String key = (String) keys.next();
                keysForSheet1.add(key.trim().toLowerCase());
            }
        }
        System.out.println("Sheet1 size: " + keysForSheet1.size());
        for (String s1 : keysForSheet1) {
            System.out.println("sheet1: " + s1);
        }

        for (int i = 0; i < sheet2data.length(); i++) {
            json = sheet2data.getJSONObject(i);
            Iterator<String> keys = json.keys();
            while (keys.hasNext()) {
                String key = (String) keys.next();
                keysForSheet2.add(key.trim().toLowerCase());
            }
        }
        System.out.println("Sheet2 size: " + keysForSheet2.size());
        for (String s2 : keysForSheet2) {
            System.out.println("sheet2: " + s2);
        }

        for (String keys : keysForSheet1) {
            System.out.println("key check: " + keys);
            if (keysForSheet2.contains(keys)) {
                System.out.println("testing: " + keysForSheet2.contains(keys));
                resultSetKeys.add(keys);
            }
        }

        for (String keys : keysForSheet2) {
            System.out.println("key check2: " + keys);
            if (keysForSheet1.contains(keys)) {
                System.out.println("testing2: " + keysForSheet1.contains(keys));
                resultSetKeys.add(keys);
            }
        }

        for (int i = 0; i < keysForSheet2.size(); i++) {
            String s = keysForSheet2.get(i);
            System.out.println("me: " + s);
            if ("name".equals(s)) {
                System.out.println("hi i came........");
            }
        }

        System.out.println("Result set: " + resultSetKeys);
        JSONArray rows = new JSONArray();
        for (String value : resultSetKeys) {
            for (int i = 0; i < sheet1data.length(); i++) {
                Iterator<String> keys = sheet1data.getJSONObject(i).keys();
                while (keys.hasNext()) {
                    String key = (String) keys.next();
                    if (value.equals(key)) {
                        rows.put(sheet1data.getJSONObject(i).getJSONArray(value));
                    }
                }
            }
        }
        writeDataToNewExcelFile(rows, colName);
        return rows;
    }

    private static void writeDataToNewExcelFile(JSONArray resultData, String colName) {
        try {
            FileWriter file = new FileWriter("C:\\jee-latest\\xslsheets\\result.csv");
            CSVWriter write = new CSVWriter(file);
            List<String[]> result = new ArrayList<>();
            for (int i = 0; i < resultData.length(); i++) {
                String[] arr = new String[resultData.length()];
                for (int j = 0; j < arr.length; j++) {
                    arr[j] = (String) resultData.getJSONArray(i).getString(j);
                }
                result.add(arr);
            }
            write.writeAll(result);
            write.close();
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        } catch (IOException ex) {
            ex.printStackTrace();
        } catch (JSONException e) {
            e.printStackTrace();
        }

    }
}
