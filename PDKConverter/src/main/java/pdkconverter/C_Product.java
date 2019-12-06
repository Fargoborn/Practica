package pdkconverter;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;

/**
 * @version $Id$
 * @since 0.1
 */
public class C_Product extends BaseTable implements TableOperations {

    /**
     * реализция доступа к таблице C_Product.
     */
    private String name;
    private String shortname;

    public C_Product(long id, String name, String shortname) {
        super(id);
        this.name = name;
        this.shortname = shortname;
    }

    public C_Product() throws SQLException, ClassNotFoundException {
        super("C_Product");
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
    public ArrayList<C_Product> read() throws SQLException {
        ArrayList<C_Product> c_productList = new ArrayList<>();
        //String maxId = "Select max(id) from C_Product";
        String selectTableSQL = "Select  id, name, shortname  from C_Product";
        Statement statement = connection.createStatement();
        try {
            // выбираем данные из БД
            ResultSet rs = statement.executeQuery(selectTableSQL);
            while (rs.next()) {
                long id = rs.getLong("id");
                String name = rs.getString("name");
                String shortname = rs.getString("shortname");
                C_Product c_product = new C_Product(id, name, shortname);
                c_productList.add(c_product);
            }
        } catch (SQLException e) {
            System.out.println(e.getMessage());
        }
        return c_productList;
    }
}
