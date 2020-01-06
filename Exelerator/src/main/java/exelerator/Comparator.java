package exelerator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.ListIterator;

/**
 * @version $Id$
 * @since 0.1
 *
 * Произвольное сравнение таблиц по столбцу.
 * Столбец ID должент иметь indx==0;
 */
public class Comparator {

    public  Comparator() throws IOException {
    }

    ExelBook exelBook = new ExelBook();
    Workbook book = exelBook.getBook();
    InputStreamReader isr = new InputStreamReader(System.in);
    BufferedReader reader  = new BufferedReader(isr);
    ArrayList<String> table_p_list = new ArrayList<>();
    ArrayList<String> table_s_list = new ArrayList<>();
    ArrayList<String> res_list = new ArrayList<>();
    int id = 0;
    String name = "";

    public void compare() throws IOException {
        System.out.println("Введите имя основной таблицы: ");
        String tab_p = reader.readLine();
        System.out.println("Введите индекс столбца: ");
        int col_p = Integer.parseInt(reader.readLine());
        System.out.println("Введите имя таблицы для сравнения: ");
        String tab_s = reader.readLine();
        System.out.println("Введите индекс столбца: ");
        int col_s = Integer.parseInt(reader.readLine());
        System.out.println("Обрабатываем: ");
        System.out.println(tab_p + "(" + col_p +")" + " /> " + tab_s + "(" + col_s +")");

        comp(tab_p, col_p, table_p_list);
        comp(tab_s, col_s, table_s_list);

        //Прямое сравнение
        for (String rec_p : table_p_list) {
            String[] strings_p = rec_p.split("@@");
            String id_p = strings_p[0];
            String name_p = strings_p[1];
            System.out.println(id_p + "  " +name_p);
            ListIterator<String> iterator = table_s_list.listIterator();
            if (iterator.hasNext()){
                while (iterator.hasNext()) {
                    String[] strings_s = iterator.next().split("@@");
                    if (id_p.equals(strings_s[0])) {
                        break;
                    } else if (!iterator.hasNext()){
                        res_list.add(rec_p);
                    }
                }
            }
        }
        for (String res : res_list) {
            System.out.println(res);
        }
    }

    private void comp(String tab, int col, ArrayList<String> list) {
        Iterator<Row> it = book.getSheet(tab).rowIterator(); // Перебираем все строки таблицы
        if (it.hasNext()) {
            it.next();
            while (it.hasNext()) {
                Row row = it.next();
                id = (int) row.getCell(0).getNumericCellValue();
                if (row.getCell(col) != null) {
                    if (row.getCell(col).getCellType() == Cell.CELL_TYPE_STRING) {
                    name = row.getCell(col).getStringCellValue().trim();
                        if (name.length() == 0) {
                        name = "----";
                        }
                    }
                    if (row.getCell(col).getCellType() == Cell.CELL_TYPE_NUMERIC) {
                        name = "" + row.getCell(col).getNumericCellValue();
                    }
                    if (row.getCell(col).getCellType() == Cell.CELL_TYPE_BLANK) {
                        name = "----";
                    }
                }
                else {
                    name = "----";
                }
                list.add(id + "@@" + name);
            }
        }
    }
}
