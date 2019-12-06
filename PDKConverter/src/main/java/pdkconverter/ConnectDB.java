package pdkconverter;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import org.postgresql.Driver;

/**
 * @version $Id$
 * @since 0.1
 */
public class ConnectDB {

    /**
     * реализция доступа к базе.
     */
    static final String DB_URL = "jdbc:postgresql://localhost:5433/sgm_bar";
    static final String USER = "postgres";
    static final String PASS = "postgres";

    ConnectDB () {

    }

    public Connection getConnectDB() {
        try {
            Class.forName("org.postgresql.Driver");
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        }
        System.out.println("Where is your MySQL JDBC Driver?");

        Connection connection = null;
        try {
            connection = DriverManager.getConnection(DB_URL, USER, PASS);
        }catch (SQLException e) {
            System.out.println("Connection Failed");
            System.out.println(e.getMessage());
        }

        if (connection != null) {
        System.out.println("Вы успешно подключились к базе");
        } else {
        System.out.println("Failed to make connection to database");
        }
        return connection;
    }
}

