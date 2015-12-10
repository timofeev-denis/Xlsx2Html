package ru.ntmedia;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
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

        //System.out.println(getDataFromXlsx("f:\\tmp\\Excel\\таблица.xlsx"));
        //return;

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
    }
    public void convertAllFiles() throws IOException {
        File dir = new File(this.srcFolder);
        if(dir == null) {
            throw new IllegalArgumentException("Не удалось открыть указанный каталог: " + this.srcFolder );
        }
        for( File f : dir.listFiles(new FilenameFilter() {
            @Override
            public boolean accept(File file, String s) {
                return s.toLowerCase().endsWith(".xlsx");
            }
        }) ) {
            convertXlsxFile(f.getCanonicalPath(), getHtmlFileName(f.getName()));
        }
    }
    public void convertXlsxFile(String srcFileName, String dstFileName) throws IOException {
        String s;
        if( (s = getDataFromXlsx(srcFileName)).equals("") ) {
            System.out.println( "EMPTY DATA" );
            return;
        }
        writeHtmlFile( dstFileName, s);
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
        OutputStreamWriter osw = new OutputStreamWriter(new FileOutputStream(fileName), "CP1251");
        osw.write(data);
        //osw.write(String.valueOf(Charset.forName("CP1251").encode(data)));
        //fw.write(String.valueOf(Charset.forName("UTF-8").encode(data)));
        osw.close();
    }
    public String getDataFromXlsx(String fileName) throws IOException {
        ArrayList<String[]> tableData = new ArrayList<>();
        String result = "";
        int rowIndex = -1;

        FileInputStream file = new FileInputStream(fileName);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            rowIndex++;
            String[] rowData = new String[TABLE_COL_COUNT];
            boolean addRow = true;
            Row row = rowIterator.next();
            for (int colIndex = 0; colIndex < TABLE_COL_COUNT; colIndex++) {
                Cell cell = row.getCell(colIndex);
                if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
                    rowData[colIndex] = "";
                } else {
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_STRING:
                            rowData[colIndex] = cell.getStringCellValue().replace("\u00A0", " ").trim();
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            rowData[colIndex] = String.valueOf(cell.getNumericCellValue());
                            break;
                    }
                }
            }
            if (rowData[0].equals("") || rowData[1].equals("")) {
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
                for (int i = 0; i < TABLE_COL_COUNT; i++) {
                    if (!rowData[i].equals("")) {
                        if (!tmp[i].equals("")) {
                            tmp[i] += "<br>";
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
        return addHtml(tableData);
    }
    public String addHtml(ArrayList<String[]> tableData) {
        String result = "<style>\n" +
                ".table-hover {\n" +
                "\tborder-spacing: 0;\n" +
                "    border-collapse: collapse;\n" +
                "}\n" +
                ".table-hover tr:hover {\n" +
                "    background-color: #f5f5f5;\n" +
                "}\n" +
                ".table-hover td {\n" +
                "    padding: 8px;\n" +
                "    border-top: 1px solid #ddd;\n" +
                "    border-bottom: 1px solid #ddd;\n" +
                "    min-width: 145px;\n" +
                "    vertical-align: top;\n" +
                "}\n" +
                "</style>\n";

        result += "<table class='table-hover'>\n";
        for (String[] s : tableData) {
            result += String.format("\t<tr><td width=30%%>%s</td><td height=70%%>%s</td></tr>\n", s[0], s[1]);
        }
        result += "</table>";
        return result;
    }
    public static String updateLastCell(String result, String data) {
        return result;
    }
}
