package ru.ntmedia;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Path;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

/**
 * Hello world!
 *
 */
public class App {
    public static void main(String[] args) {
        convertXlsxFile("file2.xlsx");
    }
    public static boolean convertXlsxFile(String fileName) {
        boolean result = false;
        String tableData = "<table>\n";
        try {
            FileInputStream file = new FileInputStream(fileName);
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            while(rowIterator.hasNext()) {
                boolean addRow = true;
                String tableRow = "\t<tr>";
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                rowloop:
                while(cellIterator.hasNext()) {
                    String tableCell = "<td>";
                    Cell cell = cellIterator.next();
                    switch(cell.getCellType()) {
                        case Cell.CELL_TYPE_STRING:
                            tableCell += cell.getStringCellValue().toString().trim();
//                            System.out.print( "[" + cell.getStringCellValue() + "]");
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
//                            System.out.print( "N: " + cell.getNumericCellValue());
                            break;
                        case Cell.CELL_TYPE_BLANK:
                            addRow = false;
                            break rowloop;
                        default:
//                            System.out.print( cell.getCellType() );
                    }
                    tableCell += "</td>";
                    tableRow += tableCell;
//                    System.out.print( "\t" );
                }
                if(addRow) {
                    tableRow += "</tr>\n";
                    tableData += tableRow;
                }
            }
            tableData += "</table>";
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
            result = false;
        }
        System.out.print(tableData);
        return result;
    }
    public static void createSpreasShit() {
        //Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Employee Data");

        //This data needs to be written (Object[])
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[]{"ID", "NAME", "LASTNAME"});
        data.put("2", new Object[]{1, "Amit", "Shukla"});
        data.put("3", new Object[]{2, "Lokesh", "Gupta"});
        data.put("4", new Object[]{3, "John", "Adwards"});
        data.put("5", new Object[]{4, "Brian", "Schultz"});

        //Iterate over data and write to sheet
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset) {
            Row row = sheet.createRow(rownum++);
            Object[] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof String)
                    cell.setCellValue((String) obj);
                else if (obj instanceof Integer)
                    cell.setCellValue((Integer) obj);
            }
        }
        try {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("howtodoinjava_demo.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("howtodoinjava_demo.xlsx written successfully on disk.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
