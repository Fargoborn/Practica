package main.java.exelerator;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class C_DocNorm {

    public  C_DocNorm() throws IOException {
            }

    SqlND sqlND = new SqlND();
    ExelBook exelBook = new ExelBook();
    Workbook book = exelBook.getBook();
    int index_id = 0;
    int index_name = 1;
    int index_shortname = 2;
    int f = 0;

    int id = 0;
    String name = "";
    String shortname = "";
    int yearlast = 0;

    ArrayList<String> sql_list = new ArrayList<String>();
    ArrayList<Integer> sgm_id = new ArrayList<Integer>();

    String sgm_sheet_name = "СГМ_НД";
    String nd__sheet_name = "Классификатор с ID без дублей";
    String result = "Результат";
    String upddel = "Удаление_перенос";
    String updfc = "ФЦ";

    public void run_C_DocNorm() throws IOException {

    Iterator<Row> it = book.getSheet(sgm_sheet_name).rowIterator(); // Перебираем все строки СГМ_НД
        if (it.hasNext()) {
        it.next();
        while (it.hasNext()) {
            Row row = it.next();
            id = (int) row.getCell(index_id).getNumericCellValue();
            //name = row.getCell(index_name).getStringCellValue();
            //shortname = row.getCell(index_shortname).getStringCellValue();
            //sgm_list.add(id + "@@" + name + "@@" + shortname + "@@");
            sgm_id.add(id);
        }
    }

    Iterator<Row> it1 = book.getSheet(nd__sheet_name).rowIterator(); // Перебираем все строки Классификатор с ID без дублей, ищем разницу с СГМ_НД
        if (it1.hasNext()) {
        it1.next();
        while (it1.hasNext()) {
            Row row = it1.next();
            id = 0;
            name = "";
            shortname = "";
            yearlast = 0;
            id = (int) row.getCell(index_id).getNumericCellValue();
            name = row.getCell(index_name).getStringCellValue();
            shortname = row.getCell(index_shortname).getStringCellValue();
            if (!sgm_id.contains(id)) {
                Row res_row = book.getSheet(result).createRow(f);
                res_row.createCell(0).setCellValue(id);
                res_row.createCell(1).setCellValue(name);
                res_row.createCell(2).setCellValue(shortname);
                sql_list.add(sqlND.insert(id, name, shortname, yearlast));
                f++;
                //System.out.println(id);
            }
        }
    }

    ArrayList <String> upddel_list = new ArrayList<String>();//список для обработки
    ArrayList <String> final_upddel_list = new ArrayList<String>();//список для скрипта
    Iterator<Row> it2 = book.getSheet(upddel).rowIterator(); // Перебираем все строки Удаление_перенос
        if (it2.hasNext()) {
        it2.next();
        while (it2.hasNext()) {
            Row row = it2.next();
            int old_id = (int) row.getCell(0).getNumericCellValue();
            int new_id = (int) row.getCell(2).getNumericCellValue();
            upddel_list.add(old_id + "&&" + new_id);
        }
    }

    ArrayList <String> updFC = new ArrayList<String>();
    Iterator<Row> it3 = book.getSheet(updfc).rowIterator(); // Перебираем все строки коррекция ФЦ
        if (it3.hasNext()) {
        it3.next();
        while (it3.hasNext()) {
            Row row = it3.next();
            id = 0;
            name = "";
            shortname = "";
            yearlast = 0;
            id = (int) row.getCell(0).getNumericCellValue();
            name = row.getCell(1).getStringCellValue();
            shortname = row.getCell(2).getStringCellValue();
            yearlast = (int) row.getCell(3).getNumericCellValue();
            updFC.add(sqlND.update(id, name, shortname, yearlast));
        }
    }

    //Готовим список для скрипта, переносы и удаления по связанным таблицам (синтаксис MSSql, Interbase, PostgreSql)
        for (String string : upddel_list) {
        String[] str = string.split("&&");
        final_upddel_list.add(sqlND.updel_Postgre(str[0], str[1]));
    }
        for (String string : upddel_list) {
        String[] str = string.split("&&");
        final_upddel_list.add(sqlND.updel_MSSQL(str[0], str[1]));
    }
        for (String string : upddel_list) {
        String[] str = string.split("&&");
        final_upddel_list.add(sqlND.updel_InterBase(str[0], str[1]));
    }
        for (String string : upddel_list) {
        String[] str = string.split("&&");
        final_upddel_list.add(sqlND.del(str[0]));
    }

    //Пишем скрипты в файлы
        System.out.println("Всего записей: " + f);
        System.out.println("******");
        System.out.println("Пишем скрипт.");
        sqlND.writesql(exelBook.url + "\\" + "_Ins_C_DocNorm", sql_list);
        sqlND.writesql(exelBook.url+ "\\"+ "_Upd_C_DocNorm", updFC);
        sqlND.writesql(exelBook.url + "\\" + "_C_DocNorm_upd_transfer.sql", final_upddel_list);
        exelBook.writebook(book, exelBook.f_url);
    }

}
