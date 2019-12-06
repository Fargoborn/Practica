package pdkconverter;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class SQLBuilder {

    public class PDKSQL {

        public PDKSQL() throws IOException {
        }

        FileInputStream inputStream = new FileInputStream("C:\\JAVA_EXEL\\PDKConverter\\normative.xls");
        Workbook doc = new HSSFWorkbook(inputStream);
        Sheet sheet = doc.getSheet("1");
        Iterator<Row> it = sheet.rowIterator();// Перебираем все строки

        public String insertPDKProduct() {
            String sql = "";
            if (it.hasNext()) {
                it.next();
                while (it.hasNext()) {
                    Row row = it.next();

                }
            }
            return sql;
        }
     }
}
