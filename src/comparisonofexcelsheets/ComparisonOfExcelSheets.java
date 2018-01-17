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
        JSONArray file1 = c.excelToJsonForSheet2(excel1, colName);
        JSONArray file2 = c.excelToJsonForSheet2(excel2, colName);
        c.comparingTwoSheets(file1, file2, colName);

    }

    public JSONArray excelToJsonForSheet2(String excelSheet, String colName) throws FileNotFoundException, IOException, InvalidFormatException, JSONException {
        CSVReader reader = new CSVReader(new FileReader(excelSheet));
        List<String[]> records = reader.readAll();
        JSONArray rows = new JSONArray();
        JSONObject jRow = new JSONObject();
        int columnNumber = 0;
        String[] a = records.get(0);
        for (int k = 0; k < a.length; k++) {
            if (colName.equals(a[k])) {
                columnNumber = k;
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
                keysForSheet1.add(key);
            }
        }

        for (int i = 0; i < sheet2data.length(); i++) {
            json = sheet2data.getJSONObject(i);
            Iterator<String> keys = json.keys();
            while (keys.hasNext()) {
                String key = (String) keys.next();
                keysForSheet2.add(key);
            }
        }

        for (String keys : keysForSheet1) {
            if (keysForSheet2.contains(keys)) {
                resultSetKeys.add(keys);
            }
        }

        for (String keys : keysForSheet2) {
            if (keysForSheet1.contains(keys)) {
                resultSetKeys.add(keys);
            }
        }

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
                        System.out.println("values: " + item.getJSONArray(value));
                    }
                }
            }
        }
        writeDataToNewExcelFile(rows, colName);
        return rows;
    }

    private static void writeDataToNewExcelFile(JSONArray resultData, String columnName) throws FileNotFoundException, JSONException {
        try {
            FileWriter file = new FileWriter("C:\\jee-latest\\xslsheets\\result.csv");
            CSVWriter write = new CSVWriter(file);
            List<String[]> result = new ArrayList<>();
            List<String[]> forHeader = new ArrayList<>();
            int arrayNumber = 0;
            for (int i1 = 0; i1 < resultData.length(); i1++) {
                String[] header = new String[resultData.getJSONArray(i1).length()];
                for (int j = 0; j < header.length; j++) {
                    if (resultData.getJSONArray(i1).getString(j).equals(columnName)) {
                        arrayNumber = i1;
                        for (int k = 0; k < header.length; k++) {
                            header[k] = (String) resultData.getJSONArray(i1).getString(k);
                        }
                        forHeader.add(header);
                    }
                }
            }
            for (int i = 0; i < resultData.length(); i++) {
                String[] arr = new String[resultData.getJSONArray(i).length()];
                for (int j = 0; j < arr.length; j++) {
                    if (i!=arrayNumber) {
                        arr[j] = (String) resultData.getJSONArray(i).getString(j);
                        System.out.println("in file " + arr[j]);
                    }
                }
                result.add(arr);
            }
            write.writeAll(forHeader);
            write.writeAll(result);
            write.close();
        } catch (FileNotFoundException ex) {
            System.out.println("Result.csv file is open please close it and try again " + ex);;
        } catch (IOException ex) {
            System.out.println("file cannot be found to close " + ex);
        }

    }
}
