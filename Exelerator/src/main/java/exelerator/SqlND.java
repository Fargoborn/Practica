package main.java.exelerator;

import java.io.BufferedWriter;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;


/**
 * @version $Id$
 * @since 0.1
 *
 * Генерация скриптов на обновление, перенос и вставку записей для C_DocNormative
 */
public class SqlND {

    public  SqlND() {
    }

    //Скрипт на вставку
    public String insert(int id, String name, String shotname, int yearlast) {
        return ("insert into C_DocNormative " + "(ID, NAME, ShortName, C_YearLast)\n" + " values (" + id + ", '" + name + "', '" + shotname +"', " + yearlast + ");\n\n");
    }

    //Скрипт на обновление
    public String update(int id, String name, String shotname, int yearlast) {
        return ("update C_DocNormative  set \n NAME = '" + name + "', \n" + " ShortName = '" + shotname + "', \n" + " C_YearLast = " + yearlast +"\n where ID = " + id + ";\n\n");
    }

    //Список переносов и удалений по связанным таблицам (синтаксис MSSql, Interbase, PostgreSql)
    public String updel_Postgre  (String old_id, String new_id) {
        return ("select TransfareID ('C_DocNormative', " + old_id + ", " + new_id + ");\n");
    }

    public String updel_MSSQL  (String old_id, String new_id) {
        return ("EXECUTE TransfareID \"C_DocNormative\", " + old_id + ", " + new_id + ");\n");
    }

    public String updel_InterBase  (String old_id, String new_id) {
        return ("EXECUTE PROCEDURE TransfareID (\"C_DocNormative\", " + old_id + ", " + new_id + ");\n");
    }

    public String del  (String old_id) {
        return ("delete from C_DocNormative where id = " + old_id + ";\n");
    }

    public void writesql (String url, ArrayList<String> list) throws IOException {
        BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(url + ".sql"), "Windows-1251"));
        for(String string : list) {
        {
            bw.write(string);
        }
        bw.flush();
        }
    }
}
