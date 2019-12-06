package pdkconverter;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.util.Iterator;

public class Separator {



    public static void main (String[] args) throws IOException, SQLException {

        FileInputStream inputStream = new FileInputStream("C:\\JAVA_EXEL\\PDKConverter\\normative.xls");
        Workbook doc = new HSSFWorkbook(inputStream);
        Sheet sheet = doc.getSheet("PDK_in");


        Iterator<Row> rowIterator = sheet.rowIterator();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                //System.out.println(row.getRowNum());
                String line = getLine(row, 0);
                System.out.println(line.trim());

                for (String l : lineCutter(line, ",")) {
                    System.out.println(l.trim());
                }
            }

            separate(0);
    }

    /**
     * Метод переноса данных из первичной таблицы
     * в таблицу c разбиением на отдельные показатели, продукты.
     * index_col - индекс столбца по которому идет разбиение.
     */
    public static void separate(int index_col) throws IOException {
        //задаем индексы для остальных столбцов
        int indx_col1 = 0;
        int indx_col2 = 0;
        if (index_col == 0) {
            indx_col1 = 1;
            indx_col2 = 2;
        }
        else if (index_col == 1) {
            indx_col1 = 0;
            indx_col2 = 2;
        }
        else {
            indx_col1 = 0;
            indx_col2 = 1;
        }

        FileInputStream inputStream = new FileInputStream("C:\\JAVA_EXEL\\PDKConverter\\normative.xls");
        Workbook doc = new HSSFWorkbook(inputStream);
        Sheet sheet = doc.getSheet("PDK_in");
        int indx = 0;
        Iterator<Row> rowIterator = sheet.rowIterator();
        String temp_line = "";

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            System.out.println(row.getRowNum());
            String line = getLine(row, index_col);
            if (!line.equals("**")) {
            temp_line = line;
            }
            String line1 = getLine(row, indx_col1);
            String line2 = getLine(row,indx_col2);
            String[] lines = lineCutter(line, ",");
            String[] temp_lines = lineCutter(temp_line, ",");

            if (!line.equals("**")) {
                for (int i = 0; i < lines.length; i++) {
                    lineWriter(indx, lines[i], line1, line2);
                    indx ++;
                }
            }
            else if (line.equals("**")) {
                System.out.println("пусто");
                System.out.print(indx + " " + line1 + "  ");
                System.out.println(line2 + " " + temp_lines.length);
                for (int i = 0; i < temp_lines.length; i++) {
                    lineWriter(indx, temp_lines[i], line1, line2);
                    indx ++;
                }
            }
        }
        }
    /**
     * Метод возвращает строку для разбиения
     */
    public static String getLine(Row row, int indx) {
        String line = "";
        Cell line_cell = row.getCell(indx);
        if (line_cell == null || line_cell.getCellType() == Cell.CELL_TYPE_BLANK) {
        line = "**";
        } else {
            line = line_cell.getStringCellValue();
        }
        return line.trim();
    }

    /**
     * Метод разбивает строку на подстроки
     * mark - разделитель(знак препинания)
     */
    public static String[] lineCutter (String line, String mark) {
        char[] p = ";".toCharArray();
        char[] lin = line.trim().toCharArray();
        int i;
        for (i = 0; i < lin.length; i++) {
            String character = String.valueOf(lin[i]);
            //System.out.println(character);
            if (character.equals("(")) {
                    for (int j = i; j < lin.length; j++) {
                        String ch = String.valueOf(lin[j]);
                        if (ch.equals(mark)) {
                            lin[j] = p[0];
                        } else if (ch.equals(")")) {
                            break;
                        }
                    }
            }
        }
        String newline = String.valueOf(lin);
        String[] list;
        list = newline.split(",");
        if (!newline.contains(",")) {
            list[0] = newline;
        }
        return list;
    }

    public static void lineWriter(int indx, String line, String line1, String line2) throws IOException {
        FileInputStream inputStream = new FileInputStream("C:\\JAVA_EXEL\\PDKConverter\\normative.xls");
        Workbook doc = new HSSFWorkbook(inputStream);
        Sheet sheetOUT = doc.getSheet("PDK_Out");

        Row row = sheetOUT.createRow(indx);
        Cell cell = row.createCell(0);
        cell.setCellValue(line);
        Cell cell1 = row.createCell(1);
        cell1.setCellValue(line1);
        Cell cell2 = row.createCell(2);
        cell2.setCellValue(line2);

        FileOutputStream fileOut = new FileOutputStream("C:\\JAVA_EXEL\\PDKConverter\\normative.xls");
        doc.write(fileOut);
        fileOut.close();
        inputStream.close();
        doc.close();
    }
}
