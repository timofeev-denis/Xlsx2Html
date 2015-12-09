package ru.ntmedia;

import freemarker.template.Configuration;
import freemarker.template.TemplateExceptionHandler;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;

/**
 * Hello world!
 *
 */
public class App {
    private final String srcFolder;
    private final String destFolder;
    private final static int TABLE_COL_COUNT = 2;

    public App(String srcFolder, String destFolder) {
        this.srcFolder = srcFolder;
        this.destFolder = destFolder;
    }

    public static void main(String[] args) {

        System.out.println(getDataFromXlsx("e:\\tmp\\Excel\\Книга10.xlsx"));
        return;
        /*
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception e) {
            e.printStackTrace();
        }

        javax.swing.SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                MainDialog dlg = new MainDialog();
                dlg.pack();
                dlg.setTitle("Конвертация файлов из Excel в HTML");
                dlg.setLocationRelativeTo(null);
                dlg.setVisible(true);
            }
        });
        */
    }
    public boolean getTemplateCfg(String templatePath) {
        Configuration cfg = new Configuration(Configuration.VERSION_2_3_22);
        try {
            cfg.setDirectoryForTemplateLoading(new File(templatePath));
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
        cfg.setDefaultEncoding("UTF-8");
        cfg.setTemplateExceptionHandler(TemplateExceptionHandler.RETHROW_HANDLER);
        return true;
    }
    public void convertAllFiles() {
        File dir = new File(this.srcFolder);
        if(dir == null) {
            return;
        }
        for( File f : dir.listFiles() ) {
            try {
                convertXlsxFile(f.getCanonicalPath(), getHtmlFileName(f.getName()));
            } catch (IOException e) {
                e.printStackTrace();
            }
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
        ArrayList<String[]> tableData = new ArrayList<>();
        String result = "";
        int rowIndex = -1;
        try {
            FileInputStream file = new FileInputStream(fileName);
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            while(rowIterator.hasNext()) {
                rowIndex++;
                String[] rowData = new String[TABLE_COL_COUNT];
                boolean addRow = true;
                Row row = rowIterator.next();
                for(int colIndex = 0; colIndex < TABLE_COL_COUNT; colIndex++) {
                    Cell cell = row.getCell(colIndex);
                    if( cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
                        rowData[colIndex] = "";
                    } else {
                        switch(cell.getCellType()) {
                            case Cell.CELL_TYPE_STRING:
                                rowData[colIndex] = cell.getStringCellValue().replace("\u00A0", " ").trim();
                                break;
                            case Cell.CELL_TYPE_NUMERIC:
                                rowData[colIndex] = String.valueOf(cell.getNumericCellValue());
                                break;
                        }
                    }
                    /*
                    boolean append = false;
                    Cell cell = row.getCell(colIndex);
                    if( cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
                        if (rowIndex != 0) {
                            append = true;
                        }
                        addRow = false;
                        continue;
                    }
                    String cellValue = " ";
                    switch(cell.getCellType()) {
                        case Cell.CELL_TYPE_STRING:
                            cellValue = cell.getStringCellValue().replace("\u00A0", " ").trim();
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            cellValue = String.valueOf(cell.getNumericCellValue());
                            break;
                    }
                    if(append) {
                        String[] tmp = tableData.get(rowIndex-1);
                        tmp[1] += "<br>\n" + cellValue;
                        tableData.set(rowIndex-1, tmp);
                    } else {
                        rowData[colIndex] = cellValue;
                    }
                    */
                }
                if(rowData[0].equals("") || rowData[1].equals("") ) {
                    // Была пустая ячейка
                    if (rowIndex == 0) {
                        // Пропускаем строку с названием
                        continue;
                    }
                    String[] tmp;
                    if (tableData.size() > 0) {
                        tmp = tableData.get(tableData.size() - 1);
                    } else {
                        tmp = new String[TABLE_COL_COUNT];
                        for (int x = 0; x < TABLE_COL_COUNT; x++) {
                            tmp[x] = "";
                        }
                    }
                    for(int i = 0; i < TABLE_COL_COUNT; i++) {
                        if(!rowData[i].equals("")) {
                            if(!tmp[i].equals("")) {
                                tmp[i] += "<br>\n";
                            }
                            tmp[i] += rowData[i];
                        }
                    }
                    if (tableData.size() > 0) {
                        tableData.set(tableData.size() - 1, tmp);
                    } else {
                        tableData.add(tmp);
                    }

                } else {
                    // Все ячейки заполнены
                    tableData.add(rowData);
                }
            }
            //result += "</table>";
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }
    public static String updateLastCell(String result, String data) {
        return result;
    }
}
