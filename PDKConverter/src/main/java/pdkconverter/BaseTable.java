package pdkconverter;

import java.io.Closeable;
import java.sql.Connection;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Objects;

/**
 * @version $Id$
 * @since 0.1
 */
// Базовый класс таблицы, имеющий ключ id
public class BaseTable implements Closeable {

    protected long id;
    Connection connection;  // JDBC-соединение для работы с таблицей
    String tableName;       // Имя таблицы

    public long getId() {
        return id;
    }

    public void setId(long id) {
        this.id = id;
    }

    public BaseTable() {}

    public BaseTable(long id) {
        this.id = id;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        BaseTable baseModel = (BaseTable) o;
        return id == baseModel.id;
    }

    @Override
    public int hashCode() {
        return Objects.hash(id);
    }

    BaseTable(String tableName) throws SQLException, ClassNotFoundException { // Для реальной таблицы передадим в конструктор её имя
        this.tableName = tableName;
        ConnectDB connectDB = new ConnectDB();
        this.connection = connectDB.getConnectDB(); // Установим соединение с СУБД для дальнейшей работы
    }

    // Закрытие
    public void close() {
        try {
            if (connection != null && !connection.isClosed())
                connection.close();
        } catch (SQLException e) {
            System.out.println("Ошибка закрытия SQL соединения!");
        }
    }

    // Выполнить SQL команду без параметров в СУБД, по завершению выдать сообщение в консоль
    void executeSqlStatement(String sql, String description) throws SQLException, ClassNotFoundException {
        reopenConnection(); // переоткрываем (если оно неактивно) соединение с СУБД
        Statement statement = connection.createStatement();  // Создаем statement для выполнения sql-команд
        statement.execute(sql); // Выполняем statement - sql команду
        statement.close();      // Закрываем statement для фиксации изменений в СУБД
        if (description != null)
            System.out.println(description);
    };

    void executeSqlStatement(String sql) throws SQLException, ClassNotFoundException {
        executeSqlStatement(sql, null);
    };


    // Активизация соединения с СУБД, если оно не активно.
    void reopenConnection() throws SQLException, ClassNotFoundException {
        if (connection == null || connection.isClosed()) {
            ConnectDB connectDB = new ConnectDB();
            connection = connectDB.getConnectDB();
        }
    }
}