package pdkconverter;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;

/**
 * @version $Id$
 * @since 0.1
 */
public class C_Note extends BaseTable implements TableOperations {

    /**
    * реализция доступа к таблице C_Note.
    */

    private String name;
    private String shortname;

    public C_Note(long id, String name, String shortname) {
        super(id);
        this.name = name;
        this.shortname = shortname;
    }

    public C_Note() throws SQLException, ClassNotFoundException {
        super("C_Note");
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getShortname() {
        return shortname;
    }

    public void setShortname(String shortname) {
        this.shortname = shortname;
    }

    @Override
    public ArrayList<C_Note> read() throws SQLException {
        ArrayList<C_Note> c_noteList = new ArrayList<>();
        //String maxId = "Select max(id) from C_Note";
        String selectTableSQL = "Select  id, name, shortname  from C_Note";
        Statement statement = connection.createStatement();
        try {
            // выбираем данные из БД
            ResultSet rs = statement.executeQuery(selectTableSQL);
            while (rs.next()) {
                long id = rs.getLong("id");
                String name = rs.getString("name");
                String shortname = rs.getString("shortname");
                C_Note c_note = new C_Note(id, name, shortname);
                c_noteList.add(c_note);
            }
        } catch (SQLException e) {
            System.out.println(e.getMessage());
        }
        return c_noteList;
    }
}
