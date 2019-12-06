package pdkconverter;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.util.Iterator;

/**
 * @version $Id$
 * @since 0.1
 */
public class IngCompare  {

    String name = "";
    String shortname = "";
    long id = 0;
    String min = "";
    String max = "";
    String nd = "";
    String ndname = "";
    String exp;

    C_Ingredient ingredienttable = new C_Ingredient();
    C_Unit unittable = new C_Unit();
    C_Product producttable = new C_Product();
    C_Note notetable = new C_Note();

    public IngCompare() throws SQLException, ClassNotFoundException {
    }

    /**
     * Метод ищет совпадения по полям Name и ShortName в таблицах классификаторов (в базе)
     * и заполняет в таблице Exel соответствующие ячейки при совпадениях.
     */
    public void ingcompare() throws SQLException, IOException {

        for (C_Ingredient c_ingredient : ingredienttable.read()) {
            id = c_ingredient.id;
            name = c_ingredient.getName();
            System.out.println(id + " " + name + " " + c_ingredient.getC_typeingredient());
        }

        //for (C_Unit c_unit : unittable.read()) {
        //    id = c_unit.id;
        //    name = c_unit.getName();
        //    System.out.println(id + " " + name);
        //}
        FileInputStream inputStream = new FileInputStream("C:\\JAVA_EXEL\\PDKConverter\\IngConverter\\normative.xls");
        Workbook doc = new HSSFWorkbook(inputStream);
        Sheet sheet = doc.getSheet("1");


        //Первичная обработка
        Iterator<Row> it = sheet.rowIterator();// Перебираем все строки
        if (it.hasNext()) {
            it.next();
            it.next();
            it.next();
            while (it.hasNext()) {
                Row row = it.next();
                System.out.println(row.getRowNum());

                //Обрабатываем показатели
                for (C_Ingredient c_ingredient : ingredienttable.read()) {
                    id = c_ingredient.id;
                    name = c_ingredient.getName().trim();
                    if (c_ingredient.getShortname() != null) {
                        shortname = c_ingredient.getShortname().trim();
                    } else {
                        shortname = "***";
                    }
                    Cell nameCell = row.getCell(3);
                    String nameEx = nameCell.getStringCellValue().trim();

                    //Прямое сравнение
                    if (nameEx.toLowerCase().equals(name.toLowerCase())) {
                        row.createCell(0);
                        row.createCell(1);
                        row.getCell(0).setCellValue(id);
                        row.getCell(1).setCellValue(name);
                    } else {
                        String[] nname = shortname.split(", ");
                        for (String nam : nname) {
                            if (nameEx.toLowerCase().equals(nam.toLowerCase())) {
                                row.createCell(0);
                                row.createCell(1);
                                row.getCell(0).setCellValue(id);
                                row.getCell(1).setCellValue(name);
                            }
                        }
                    }
                    String[] nname = name.trim().split(" ");
                    if (nname.length > 1) {
                        name = nname[0] + " " + nname[1];
                        if (nameEx.toLowerCase().contains(name.toLowerCase())) {
                            row.createCell(0);
                            row.createCell(1);
                            row.getCell(0).setCellValue(id);
                            row.getCell(1).setCellValue(c_ingredient.getName().trim());
                        }
                        if (name.toLowerCase().contains(nameEx.toLowerCase())) {
                            row.createCell(0);
                            row.createCell(1);
                            row.getCell(0).setCellValue(id);
                            row.getCell(1).setCellValue(c_ingredient.getName().trim());
                        }
                    }

                    //Сравнение с перестановкой
                    if (!nameEx.toLowerCase().equals(name.toLowerCase())) {
                        if (nname.length > 1) {
                        name = nname[1] + " " + nname[0];
                        if (nameEx.toLowerCase().contains(name.toLowerCase())) {
                            if (id > 0) {
                                row.createCell(0);
                            row.getCell(0).setCellValue(id);
                            } else {
                                row.createCell(0);
                                row.getCell(0).setCellValue(0);
                            }
                            row.createCell(1);
                            row.getCell(1).setCellValue(c_ingredient.getName().trim());
                        }
                        }

                    row.createCell(4);
                    row.getCell(4).setCellValue("*");
                }
            }
        }
        inputStream.close();
        FileOutputStream fileOut = new FileOutputStream("C:\\JAVA_EXEL\\PDKConverter\\IngConverter\\normative.xls");

        doc.write(fileOut);
        fileOut.close();
        doc.close();
        inputStream.close();

        ingredienttable.close();
        unittable.close();
    }
    }

    public static void main (String[] args) throws IOException, SQLException, ClassNotFoundException {

        IngCompare ingCompare = new IngCompare();
        ingCompare.ingcompare();
    }
}
