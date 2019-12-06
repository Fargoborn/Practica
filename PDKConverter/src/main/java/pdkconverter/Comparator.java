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
public class Comparator  {

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

    public Comparator() throws SQLException, ClassNotFoundException {
    }

    /**
     * Метод ищет совпадения по полям Name и ShortName в таблицах классификаторов (в базе)
     * и заполняет в таблице Exel соответствующие ячейки при совпадениях.
     */
    public void compare() throws SQLException, IOException {

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
    FileInputStream inputStream = new FileInputStream("C:\\JAVA_EXEL\\PDKConverter\\normative.xls");
    Workbook doc = new HSSFWorkbook(inputStream);
    Sheet sheet = doc.getSheet("1");

    //Пишем код и назавние НД
    Row rownd = sheet.getRow(1);
    rownd.getCell(0).setCellValue(nd);
    rownd.getCell(1).setCellValue(ndname);

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
                Cell nameCell = row.getCell(6);
                String nameEx = nameCell.getStringCellValue().trim();
                if (nameEx.toLowerCase().equals(name.toLowerCase())) {
                    row.getCell(4).setCellValue(id);
                    row.getCell(5).setCellValue(name);
                } else {
                    String[] nname = shortname.split(", ");
                    for (String nam : nname) {
                        if (nameEx.toLowerCase().equals(nam.toLowerCase())) {
                            row.getCell(4).setCellValue(id);
                            row.getCell(5).setCellValue(name);
                        }
                    }
                }
            }
            //Обрабатываем единицы измерения
            for (C_Unit c_unit : unittable.read()) {
                id = c_unit.id;
                name = c_unit.getName().trim();
                if (c_unit.getShortname() != null) {
                    shortname = c_unit.getShortname().trim();
                }
                Cell nameCell = row.getCell(9);
                String nameEx = "";
                if (nameCell == null || nameCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                    nameCell = row.getCell(10);
                    if (nameCell == null || nameCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                        System.out.println(row.getRowNum() + "нет норматива");
                        break;
                    }
                    nameEx = nameCell.getStringCellValue().trim();
                    String[] nEx =nameEx.split(" ");
                    nameEx = nEx[1];
                    nameCell = row.getCell(9);
                    nameCell.setCellValue(nameEx);
                } else {
                    nameEx = nameCell.getStringCellValue().trim();
                }

                System.out.println(row.getRowNum() + " " + nameEx);

                if (nameEx.toLowerCase().equals(name.toLowerCase())) {
                    row.getCell(7).setCellValue(id);
                    row.getCell(8).setCellValue(name);
                } else if (nameEx.toLowerCase().equals(shortname.toLowerCase())) {
                    row.getCell(7).setCellValue(id);
                    row.getCell(8).setCellValue(name);
                } else if(nameEx.equals("")) {
                    row.getCell(7).setCellValue(0);
                    row.getCell(8).setCellValue("не найдено");
                }
            }

            //Обрабатываем продукты
            for (C_Product c_product : producttable.read()) {
                id = c_product.id;
                name = c_product.getName().trim();
                if (c_product.getShortname() != null) {
                    shortname = c_product.getShortname().trim();
                }
                Cell nameCell = row.getCell(1);
                if (nameCell == null || nameCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                    System.out.println("Нет продукта");
                    break;
                }
                String nameEx = nameCell.getStringCellValue().trim();
                if (nameEx.toLowerCase().equals(name.toLowerCase())) {
                    row.getCell(0).setCellValue(id);
                } else if (nameEx.toLowerCase().equals(shortname.toLowerCase())) {
                    row.getCell(0).setCellValue(id);
                } else if (String.valueOf(row.getCell(0).getNumericCellValue()) == null){
                    row.getCell(0).setCellValue(0);
                }
            }

            //Обрабатываем нормативы
            Cell pdkCell = row.getCell(10);
            if (pdkCell == null || pdkCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                System.out.println("Нет норматива");
                break;
            }
            String pdk = pdkCell.toString();
            if (pdk.contains("-")) {
                String[] min_max = pdk.split("-");
                min = min_max[0].trim();
                max = min_max[1].trim();
                row.getCell(11).setCellValue(min);
                row.getCell(12).setCellValue(max);
            } else if (pdk.contains("х")) {
                String pdkm[] = pdk.split("х");
                exp = pdkm[1].substring(pdkm.length - 1).replace("0", "");
                int e = Integer.parseInt(exp);
                double m = Double.parseDouble(pdkm[0].replace(",", ".")) * 10;
                for (int j = 1; j < e; j++) {
                    m = m * 10;
                }
                max = String.valueOf(m);
                row.getCell(12).setCellValue(max);
                row.getCell(13).setCellValue(exp);
            } else if(!pdk.contains("-") && !pdk.contains("х")) {
                max = pdk;
                row.getCell(12).setCellValue(max);
            }
        }
    }
    inputStream.close();
    FileOutputStream fileOut = new FileOutputStream("C:\\JAVA_EXEL\\PDKConverter\\normative.xls");

            doc.write(fileOut);
            fileOut.close();
            doc.close();
            inputStream.close();

        ingredienttable.close();
        unittable.close();
    }
}
