package exelerator;


import java.io.BufferedWriter;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;

/**
 * @version $Id$
 * @since 0.1
 *
 * Генерация скриптов на обновление, перенос и вставку записей для C_Ingredient
 */
public class SqlIngredient {

    //Скрипт на вставку
    public String insert(int id, String name, String shotname, int C_TYPEINGREDIENT, String CAS) {
        return ("insert into C_Ingredient " + "(ID, NAME, SHORTNAME, C_TYPEINGREDIENT, CAS)\n" + " values (" + id + ", '" + name + "', '" + shotname +"', " + C_TYPEINGREDIENT +
                ", '" + CAS + "');\n\n");
    }

    //Скрипт на обновление
    public String update(int id, String name, String shotname, int C_TYPEINGREDIENT, String CAS) {
        return ("update C_Ingredient  set \n NAME = '" + name + "', \n" + " ShortName = '" + shotname + "', \n" + "C_TYPEINGREDIENT = " + C_TYPEINGREDIENT +
                ", \n" + "CAS = '" + CAS + "',\n where ID = " + id + ";\n\n");
    }

    //Список переносов и удалений по связанным таблицам (синтаксис MSSql, Interbase, PostgreSql)
    public String updel_Postgre  (String old_id, String new_id) {
        return ("select TransfareID ('C_Ingredient', " + old_id + ", " + new_id + ");\n");
    }

    public String updel_MSSQL  (String old_id, String new_id) {
        return ("EXECUTE TransfareID \"C_Ingredient\", " + old_id + ", " + new_id + ";\n");
    }

    public String updel_InterBase  (String old_id, String new_id) {
        return ("EXECUTE PROCEDURE TransfareID (\"C_Ingredient\", " + old_id + ", " + new_id + ");\n");
    }

    public String del  (String old_id) {
        return ("delete from C_Ingredient where id = " + old_id + ";\n");
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
