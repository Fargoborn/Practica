package pdkconverter;

import java.sql.SQLException;
import java.util.ArrayList;

/**
 * @version $Id$
 * @since 0.1
 */
// Операции с таблицами
public interface TableOperations {
    //void createTable() throws SQLException; // создание таблицы
    //void createForeignKeys() throws SQLException; // создание связей между таблицами
    //void createExtraConstraints() throws SQLException; // создание дополнительных правил для значений полей таблиц
    public ArrayList read() throws SQLException; //чтение
    //void insert() throws  SQLException; //вставка
    //void update() throws SQLException; //обновление
}