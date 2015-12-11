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
    private enum ROW_TYPE {TITLE, SUBTITLE, TEXT};

    public App(String srcFolder, String destFolder) {
        this.srcFolder = srcFolder;
        this.destFolder = destFolder;
    }

    public static void main(String[] args) {
        /*
        App app = new App("", "");
        try {
            System.out.println(app.getDataFromXlsx("f:\\tmp\\Excel\\Книга10.xlsx"));
        } catch (IOException e) {
            e.printStackTrace();
        }
        return;
        */

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
        ArrayList<RowData> tableData = new ArrayList<>();
        String header = "";
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
            /*
            незаполненная 2-я ячейка в первой строке - это заголовок
            незаполненная 2-я ячейка во второй и далее строке - это подзаголовок
            незаполненная 1-я ячейка - это продолжение предыдущей строки
                либо новая строка
            */
            if (rowData[0].equals("") && rowData[1].equals("")) {
                // Обе ячейки пусты - пропускаем строку
                // TODO: переделать на произвольное количество ячеек в таблице
                continue;
            }
            if (rowData[0].equals("")) {
                RowData tmp = new RowData();
                if (rowIndex == 0) {
                    // Новая строка
                    tmp.data[0] = "";
                    tmp.data[1] = rowData[1];
                    tmp.rowType = ROW_TYPE.TEXT;
                    tableData.add(tmp);
                } else {
                    // Продолжение предыдущей строки
                    tmp = tableData.get(tableData.size() - 1);
                    for (int i = 0; i < TABLE_COL_COUNT; i++) {
                        if (!rowData[i].equals("")) {
                            if (!tmp.data[i].equals("")) {
                                tmp.data[i] += "<br>";
                            }
                            tmp.data[i] += rowData[i];
                        }
                    }
                    if (tableData.size() > 0) {
                        tableData.set(tableData.size() - 1, tmp);
                    } else {
                        tableData.add(tmp);
                    }
                }
            } else if(rowData[1].equals("")) {
                // Незаполненная 2-я ячейка
                RowData tmp = new RowData();
                tmp.data = rowData;
                if (rowIndex == 0) {
                    // В первой строке - это заголовок
                    tmp.rowType = ROW_TYPE.TITLE;
                } else {
                    // Во второй и далее строке - это подзаголовок
                    tmp.rowType = ROW_TYPE.SUBTITLE;
                }
                tableData.add(tmp);
            } else {
                // Все ячейки заполнены
                tableData.add(new RowData(rowData, ROW_TYPE.TEXT));
            }
        }
        return addHtml(tableData);
    }
    public String addHtml(ArrayList<RowData> tableData) {
        String result = "";

        if (tableData.size() == 0) {
            return result;
        }
        RowData r = tableData.get(0);
        int rowIndex = 0;
        if( r.rowType == ROW_TYPE.TITLE) {
            result += String.format( "<h2 class='table-hover-title'>%s</h2>\n", r.data[0]);
            rowIndex++;
        }
        result += "<table class='table-hover'>\n";
        for (; rowIndex < tableData.size(); rowIndex++) {
            r = tableData.get(rowIndex);
            switch (r.rowType) {
                case SUBTITLE:
                    result += String.format("\t<tr><td colspan=2 class='table-hover-subtitle'>%s</td></tr>\n", r.data[0], r.data[1]);
                    break;
                default:
                    result += String.format("\t<tr><td>%s</td><td>%s</td></tr>\n", r.data[0], r.data[1]);
            }

        }
        result += "</table>";
        return result;
    }
    public static String updateLastCell(String result, String data) {
        return result;
    }

    private final class RowData {
        public String[] data;
        public ROW_TYPE rowType;
        public RowData() {

        }
        public RowData(String[] data, ROW_TYPE rowType) {
            this.data = data;
            this.rowType = rowType;
        }
    }
}
