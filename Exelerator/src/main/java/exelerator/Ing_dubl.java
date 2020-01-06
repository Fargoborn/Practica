package exelerator;

import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.util.Iterator;

public class Ing_dubl {
    public  Ing_dubl() throws IOException {
    }

    ExelBook exelBook = new ExelBook();
    Workbook book = exelBook.getBook();


    int mindex_id = 0;
    int index_id = 0;
    String sheet_name = "Лист1";

    public void run() throws IOException {

        CellStyle style = book.createCellStyle();
        style.setFillBackgroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Sheet sheet = book.getSheet(sheet_name);

        Iterator<Row> it = book.getSheet(sheet_name).rowIterator(); // Перебираем все строки
        if (it.hasNext()) {
            it.next();
            while (it.hasNext()) {
                Row row = it.next();
                int indx_row = row.getRowNum();
                System.out.println(indx_row);
                Row r = sheet.getRow(indx_row + 1);
                mindex_id = (int) row.getCell(6).getNumericCellValue();
                index_id = (int) row.getCell(4).getNumericCellValue();

            if (r.getCell(4).getCellType() == Cell.CELL_TYPE_BLANK){
                    break;
                }
                if (index_id == mindex_id && mindex_id == (int) r.getCell(6).getNumericCellValue()) {
                    row.getCell(2).setCellStyle(style);
                }

            }
        }
        exelBook.writebook(book, exelBook.f_url);
    }
}
