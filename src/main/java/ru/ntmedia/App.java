package ru.ntmedia;

import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.*;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import static java.nio.charset.StandardCharsets.UTF_8;

/**
 * Hello world!
 */
public class App {
    private final String srcFolder;
    private final String destFolder;
    private final Map<ROW_TYPE, String> templates = new HashMap<>();
    private final static int TABLE_COL_COUNT = 2;

    private enum ROW_TYPE {
        SUBTITLE,
        TEXT
    }

    App(String srcFolder, String destFolder) {
        this.srcFolder = srcFolder;
        this.destFolder = destFolder;
        templates.put(ROW_TYPE.TEXT, readTemplateFromFile("/templates/data-cell.html"));
        templates.put(ROW_TYPE.SUBTITLE, readTemplateFromFile("/templates/subtitle.html"));
    }

    private static void writeHtmlFile(String fileName, String data) throws IOException {
        if (fileName == null || fileName.equals("")) {
            throw new IllegalArgumentException("Имя файла не указано.");
        }
        OutputStreamWriter osw = new OutputStreamWriter(new FileOutputStream(fileName), UTF_8);
        osw.write(data);
        //osw.write(String.valueOf(Charset.forName("CP1251").encode(data)));
        //fw.write(String.valueOf(Charset.forName("UTF-8").encode(data)));
        osw.close();
    }

    public static void main(String[] args) {
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

    private String readTemplateFromFile(String s) {
        InputStream is = App.class.getResourceAsStream(s);
        try {
            return IOUtils.toString(is, UTF_8);
        } catch (IOException e) {
            e.printStackTrace();
            return "ошибка чтения шаблона " + s;
        }
    }

    void convertAllFiles() throws IOException {
        File dir = new File(this.srcFolder);
        if (dir == null) {
            throw new IllegalArgumentException("Не удалось открыть указанный каталог: " + this.srcFolder);
        }
        for (File f : dir.listFiles(new FilenameFilter() {
            @Override
            public boolean accept(File file, String s) {
                return s.toLowerCase().endsWith(".xlsx");
            }
        })) {
            convertXlsxFile(f.getCanonicalPath(), getHtmlFileName(f.getName()));
        }
    }

    private void convertXlsxFile(String srcFileName, String dstFileName) throws IOException {
        String s;
        if ((s = getDataFromXlsx(srcFileName)).equals("")) {
            System.out.println("EMPTY DATA");
            return;
        }
        writeHtmlFile(dstFileName, s);
    }

    private String getHtmlFileName(String fileName) {
        // cut path
        String baseFileName = Paths.get(fileName).getFileName().toString();
        // cut extension
        String rootFileName = baseFileName.substring(0, baseFileName.lastIndexOf("."));
        return this.destFolder + File.separator + rootFileName + ".html";
    }

    private String getDataFromXlsx(String fileName) throws IOException {
        ArrayList<RowData> tableData = new ArrayList<>();
        int rowIndex = -1;

        FileInputStream file = new FileInputStream(fileName);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);
        for (Row aSheet : sheet) {
            rowIndex++;
            String[] rowData = new String[TABLE_COL_COUNT];
            Row row = aSheet;
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
                            double numericCellValue = cell.getNumericCellValue();
                            if (numericCellValue - (long) numericCellValue != 0) {
                                rowData[colIndex] = String.valueOf(numericCellValue);
                            } else {
                                rowData[colIndex] = String.format("%.0f", numericCellValue);
                            }
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
                                tmp.data[i] += "<br />";
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
            } else if (rowData[1].equals("")) {
                // Незаполненная 2-я ячейка
                RowData tmp = new RowData();
                tmp.data = rowData;
                tmp.rowType = ROW_TYPE.SUBTITLE;
                tableData.add(tmp);
            } else {
                // Все ячейки заполнены
                tableData.add(new RowData(rowData, ROW_TYPE.TEXT));
            }
        }
        return addHtml(tableData);
    }

    private String addHtml(ArrayList<RowData> tableData) {
        StringBuilder result = new StringBuilder();

        if (tableData.size() == 0) {
            return result.toString();
        }
        RowData r;
        int rowIndex = 0;
        for (; rowIndex < tableData.size(); rowIndex++) {
            r = tableData.get(rowIndex);
            result.append(String.format(templates.get(r.rowType), r.data[0], r.data[1]));
        }
        return result.toString();
    }

    private final class RowData {
        String[] data;
        ROW_TYPE rowType;

        RowData() {

        }

        RowData(String[] data, ROW_TYPE rowType) {
            this.data = data;
            this.rowType = rowType;
        }
    }
}
