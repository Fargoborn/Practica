package prikazall;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.Iterator;

public class Main {

    public static void main(String[] args)throws Exception {

        String Name = "";
        String Order = "";
        String Order_all = "";
        String Tabel = "";

        ExelBook exelBook = new ExelBook();
        Workbook book = exelBook.getBook();
        Sheet sheet = exelBook.getSheet(book, "Лист1");

        ArrayList<String> list = new ArrayList<>();
        ArrayList<String> fin_list = new ArrayList<>();
        Iterator<Row> it = sheet.rowIterator(); // Перебираем все строки
        if (it.hasNext()) {
            it.next();
            while (it.hasNext()) {
                Row row = it.next();
                Name = row.getCell(3).getStringCellValue().trim();
                Order = (int) row.getCell(1).getNumericCellValue() + "";
                Tabel = (int) row.getCell(2).getNumericCellValue() + "";
                list.add(Name + "&&" + Tabel + "&&" + Order);
            }
        }

        String old_tabel = "";
        Sheet res_sheet = exelBook.getSheet(book, "Лист2");
        int f = 0;
        for (int i = 0; i < list.size(); i++) {
            String[] s_strings = list.get(i).split("&&");
            Name = s_strings[0];
            Tabel = s_strings[1];
            Order = s_strings[2];
            Order_all = Order;
            if (!old_tabel.contains(Tabel)) {
            for (String string : list) {
                String[] strings = string.split("&&");
                String n = strings[0];
                String t = strings[1];
                String o = strings[2];
                if (n.equals(Name) && !o.equals(Order)) {
                    Order_all = Order_all + ", " + o;
                }
            }
            Row row = res_sheet.createRow(f + 1);
            Cell N_cell = row.createCell(0);
            N_cell.setCellValue(f + 1);
            Cell Name_cell = row.createCell(1);
            Name_cell.setCellValue(Name);
            Cell Orders_cell = row.createCell(2);
            Orders_cell.setCellValue(Order_all);
            old_tabel = old_tabel + ", " + Tabel;
            f++;
        }
        }

        exelBook.writebook(book, exelBook.f_url);
    }
}
