package pdkconverter;


import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;

/**
 * @version $Id$
 * @since 0.1
 */
public class C_Ingredient extends BaseTable implements TableOperations{

    /**
     * реализция доступа к таблице C_Ingredirnt.
     */
    private String name;
    private String shortname;
    private long c_typeingredient;
    private String cas;

    public C_Ingredient(long id, String name, String shortname, long c_typeingredient, String cas) {
        super(id);
        this.name = name;
        this.shortname = shortname;
        this.c_typeingredient = c_typeingredient;
        this.cas = cas;
    }

    public C_Ingredient() throws SQLException, ClassNotFoundException {
        super("C_Ingredient");
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

    public long getC_typeingredient() {
        return c_typeingredient;
    }

    public void setC_typeingredient(long c_typeingredient) {
        this.c_typeingredient = c_typeingredient;
    }

    public String getCas() {
        return cas;
    }

    public void setCas(String cas) {
        this.cas = cas;
    }


    @Override
    public ArrayList<C_Ingredient> read() throws SQLException {
        ArrayList<C_Ingredient> c_ingredientList = new ArrayList<>();

        //String maxId = "Select max(id) from C_Ingredient";
        String selectTableSQL = "Select  id, name, shortname, c_typeingredient, cas  from C_Ingredient";
        Statement statement = connection.createStatement();
        try {
            // выбираем данные из БД
            ResultSet rs = statement.executeQuery(selectTableSQL);
            while (rs.next()) {
                long id = rs.getLong("id");
                String name = rs.getString("name");
                String shortname = rs.getString("shortname");
                long c_typeingredient = rs.getLong("c_typeingredient");
                String cas = rs.getString("cas");
                C_Ingredient c_ingredient = new C_Ingredient(id, name, shortname, c_typeingredient, cas);
                c_ingredientList.add(c_ingredient);
            }
        } catch (SQLException e) {
            System.out.println(e.getMessage());
        }
        return c_ingredientList;
    }
}
