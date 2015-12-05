package ru.ntmedia;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Paths;
import java.util.*;

/**
 * Hello world!
 *
 */
public class App {
    private final String srcFolder;
    private final String destFolder;

    public App(String srcFolder, String destFolder) {
        this.srcFolder = srcFolder;
        this.destFolder = destFolder;
    }

    public static void main(String[] args) {
        // TODO: 05.12.2015
        App Converter = new App("f:\\tmp\\Excel", "f:\\tmp\\HTML");
        try {
            Converter.convertAllFiles();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    public void convertAllFiles() throws IOException {
        File dir = new File(this.srcFolder);
        for( File f : dir.listFiles() ) {
            convertXlsxFile(f.getCanonicalPath(), getHtmlFileName(f.getName()));
        }
    }


    public static void convertXlsxFile(String srcFileName, String dstFileName) {
        String s;
        if( (s = getDataFromXlsx(srcFileName)).equals("") ) {
            System.out.println( "EMPTY DATA" );
            return;
        }
        try {
            writeHtmlFile( dstFileName, s);
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }
    public String getHtmlFileName(String fileName) {
        // cut path
        String baseFileName = Paths.get(fileName).getFileName().toString();
        // cut extension
        String rootFileName = baseFileName.substring(0, baseFileName.lastIndexOf("."));
        return this.destFolder + File.separator + rootFileName + ".html";
    }
    public static void writeHtmlFile(String fileName, String data) throws IOException {
        if(fileName == null || fileName.equals("")) {
            throw new IllegalArgumentException( "Имя файла не указано." );
        }
        FileWriter fw = new FileWriter(fileName);
        fw.write(data);
        fw.close();
    }
    public static String getDataFromXlsx(String fileName) {
        String result = "<table>\n";
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
                            tableCell += cell.getStringCellValue().replace("\u00A0", " ").trim();
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            break;
                        case Cell.CELL_TYPE_BLANK:
                            addRow = false;
                            break rowloop;
                        default:
                    }
                    tableCell += "</td>";
                    tableRow += tableCell;
                }
                if(addRow) {
                    tableRow += "</tr>\n";
                    result += tableRow;
                }
            }
            result += "</table>";
        } catch (Exception e) {
            e.printStackTrace();
            result = "";
        }
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
