package pdkconverter;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;

public class C_Unit extends BaseTable implements TableOperations {

    /**
     * реализция доступа к таблице C_Unit.
     */

    private String name;
    private String shortname;

    public C_Unit(long id, String name, String shortname) {
        super(id);
        this.name = name;
        this.shortname = shortname;
    }

    public C_Unit() throws SQLException, ClassNotFoundException {
        super("C_Unit");
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
    public ArrayList<C_Unit> read() throws SQLException {
        ArrayList<C_Unit> c_unitList = new ArrayList<>();
        //String maxId = "Select max(id) from C_Unit";
        String selectTableSQL = "Select  id, name, shortname  from C_Unit";
        Statement statement = connection.createStatement();
        try {
            // выбираем данные из БД
            ResultSet rs = statement.executeQuery(selectTableSQL);
            while (rs.next()) {
                long id = rs.getLong("id");
                String name = rs.getString("name");
                String shortname = rs.getString("shortname");
                C_Unit c_unit = new C_Unit(id, name, shortname);
                c_unitList.add(c_unit);
            }
        } catch (SQLException e) {
            System.out.println(e.getMessage());
        }

        return c_unitList;
    }
}
