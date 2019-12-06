package main.java.exelerator;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class C_Ingredient {

    public  C_Ingredient() throws IOException {
    }

    SqlIngredient sqlIngredient = new SqlIngredient();
    ExelBook exelBook = new ExelBook();
    Workbook book = exelBook.getBook();
    int index_id = 0;
    int index_name = 1;
    int index_shortname = 2;
    int index_typeingredient = 3;
    int index_cas = 4;
    int f = 0;

    int id = 0;
    String name = "";
    String shortname = "";
    int typeingredient = 0;
    String cas = "";

    String ing_sheet = "Показатель";
    String trans_sheet = "Перенос";

    ArrayList<String> ins_list = new ArrayList<String>();
    ArrayList <String> upd_list = new ArrayList<String>();
    ArrayList <String> trans_list = new ArrayList<String>();

    public void run_C_Ingredient() throws IOException {

        Iterator<Row> it = book.getSheet(ing_sheet).rowIterator(); // Перебираем все строки Показатель
        if (it.hasNext()) {
            it.next();
            while (it.hasNext()) {
                Row row = it.next();
                id = (int) row.getCell(index_id).getNumericCellValue();
                name = row.getCell(index_name).getStringCellValue();
                shortname = row.getCell(index_shortname).getStringCellValue();
                typeingredient = (int) row.getCell(index_typeingredient).getNumericCellValue();
                cas = row.getCell(index_cas).getStringCellValue();

                ins_list.add(sqlIngredient.insert(id, name, shortname, typeingredient, cas));
                upd_list.add(sqlIngredient.update(id, name, shortname, typeingredient, cas));
                f++;
            }
        }

        ArrayList <String> upd_del_list = new ArrayList<String>();
        Iterator<Row> it1 = book.getSheet(trans_sheet).rowIterator(); // Перебираем все строки Перенос
        if (it1.hasNext()) {
            it1.next();
            while (it1.hasNext()) {
                Row row = it1.next();
                int old_id = (int) row.getCell(0).getNumericCellValue();
                int new_id = (int) row.getCell(2).getNumericCellValue();

                upd_del_list.add(old_id + "&&" + new_id);
            }
        }

        //Готовим список для скрипта, переносы и удаления по связанным таблицам (синтаксис MSSql, Interbase, PostgreSql)
        for (String string : upd_del_list) {
            String[] str = string.split("&&");
            trans_list.add(sqlIngredient.updel_Postgre(str[0], str[1]));
        }
        for (String string : upd_del_list) {
            String[] str = string.split("&&");
            trans_list.add(sqlIngredient.updel_MSSQL(str[0], str[1]));
        }
        for (String string : upd_del_list) {
            String[] str = string.split("&&");
            trans_list.add(sqlIngredient.updel_InterBase(str[0], str[1]));
        }
        for (String string : upd_del_list) {
            String[] str = string.split("&&");
            trans_list.add(sqlIngredient.del(str[0]));
        }

        //Пишем скрипты в файлы
        System.out.println("Всего записей: " + f);
        System.out.println("******");
        System.out.println("Пишем скрипт.");
        sqlIngredient.writesql(exelBook.url + "\\" + "_Ins_C_Ingredient", ins_list);
        sqlIngredient.writesql(exelBook.url+ "\\"+ "_Upd_C_Ingredient", upd_list);
        sqlIngredient.writesql(exelBook.url + "\\" + "_C_Ingredient_upd_transfer.sql", trans_list);
    }
}
