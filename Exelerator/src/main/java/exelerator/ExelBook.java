package exelerator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;

/**
 * @version $Id$
 * @since 0.1
 */
public class ExelBook {
    public String f_url = "C:\\Users\\dbarshhevskijj\\Desktop\\132887\\без дублей_Классификатор_НД.xls";
    public String url = "";
    String sheet_name = "";


    public ExelBook() {
    }

    //получаем книгу
    public Workbook getBook() throws IOException {
        System.out.println("Введите путь к файлу: ");
        InputStreamReader isr = new InputStreamReader(System.in);
        BufferedReader reader  = new BufferedReader(isr);
        url = reader.readLine();
        System.out.println("Введите имя файла: ");
        f_url = url + "\\" + reader.readLine();
        System.out.println(f_url);

        FileInputStream inputStream = new FileInputStream(f_url);
        return new HSSFWorkbook(inputStream);
    }

    //пишем книгу
    public void writebook(Workbook book, String f_url) throws IOException {
        FileOutputStream fileOut = new FileOutputStream(f_url);
        book.write(fileOut);
        fileOut.close();
        book.close();
    }

    public Sheet getSheet(Workbook book, String sheet_name) {
        return book.getSheet(sheet_name);
    }

    public Row getRow(Sheet sheet, int index) {
        return sheet.getRow(index);
    }

    public Cell getCell(Row row, int cell_index) {
        return row.getCell(cell_index);
    }
}
