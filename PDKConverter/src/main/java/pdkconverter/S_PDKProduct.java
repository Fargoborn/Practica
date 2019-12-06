package pdkconverter;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;

public class S_PDKProduct extends BaseTable {

    /**
     * реализция доступа к таблице S_PDKProduct.
     */
    public S_PDKProduct(long id) {
        super(id);
    }

    public S_PDKProduct() throws SQLException, ClassNotFoundException {
        super("S_PDKProduct");
    }

    public long getMaxID() throws SQLException {
        ArrayList<S_PDKProduct> s_pdkProducts = new ArrayList<>();
        String maxId = "Select max(id) from S_PDKProduct";
        Statement statement = connection.createStatement();
        try {
            // выбираем данные из БД
            ResultSet rs = statement.executeQuery(maxId);
            while (rs.next()) {
                long id = rs.getLong("MAX");
                S_PDKProduct s_pdkProduct = new S_PDKProduct(id);
                s_pdkProducts.add(s_pdkProduct);
            }
        } catch (SQLException e) {
            System.out.println(e.getMessage());
        }
        return s_pdkProducts.get(0).getId();
    }
}
