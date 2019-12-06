package pdkconverter;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;

public class C_Docnormative extends BaseTable implements TableOperations{

    private String name;
    private String shortname;
    private int c_yearlast;
    //private int foodguality;


    public C_Docnormative(long id, String name, String shortname, int c_yearlast) {
        super(id);
        this.name = name;
        this.shortname = shortname;
        this.c_yearlast = c_yearlast;
        //this.foodguality = foodguality;
    }

    public C_Docnormative() throws SQLException, ClassNotFoundException {
        super("C_Docnormative");
    }

    public String getName() {
        return name;
    }

    public String getShortname() {
        return shortname;
    }

    public int getC_yearlast() {
        return c_yearlast;
    }



    public void setName(String name) {
        this.name = name;
    }

    public void setShortname(String shortname) {
        this.shortname = shortname;
    }

    public void setC_yearlast(int c_yearlast) {
        this.c_yearlast = c_yearlast;
    }




    @Override
    public ArrayList<C_Docnormative> read() throws SQLException {
        ArrayList<C_Docnormative> c_docnormativeslist = new ArrayList<>();

        //String maxId = "Select max(id) from C_Ingredient";
        String selectTableSQL = "Select  id, name, shortname, c_yearlast  from C_Docnormative";
        Statement statement = connection.createStatement();
        try {
            // выбираем данные из БД
            ResultSet rs = statement.executeQuery(selectTableSQL);
            while (rs.next()) {
                long id = rs.getLong("id");
                String name = rs.getString("name");
                String shortname = rs.getString("shortname");
                int c_yearlast = rs.getInt("c_yearlast");
                //int foodguality = rs.getInt("cas");
                C_Docnormative c_docnormative = new C_Docnormative(id, name, shortname, c_yearlast);
                c_docnormativeslist.add(c_docnormative);
            }
        } catch (SQLException e) {
            System.out.println(e.getMessage());
        }
        return c_docnormativeslist;
    }


}
