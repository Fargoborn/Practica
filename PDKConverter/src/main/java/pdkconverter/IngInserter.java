package pdkconverter;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Iterator;

public class IngInserter {

    C_Ingredient ingredienttable = new C_Ingredient();

    public IngInserter() throws SQLException, ClassNotFoundException, IOException {
    }

    private class Ingredient {
        String name = "";
        String indx = "";
        private Ingredient(String name, String indx) {
            this.name = name;
            this.indx = indx;
        }

        public String getName() {
            return name;
        }

        public String getIndx() {
            return indx;
        }
    }

    public ArrayList<Ingredient> getIngredients() throws IOException {
        ArrayList<Ingredient> list = new ArrayList<>();
        FileInputStream inputStream = new FileInputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\Новая папка\\normative.xls");
        Workbook doc = new HSSFWorkbook(inputStream);
        Sheet sheet = doc.getSheet("1");
        Iterator<Row> it = sheet.rowIterator();
        if (it.hasNext()) {
            while (it.hasNext()) {
                Row row = it.next();
                Cell cell_name = row.getCell(0);
                Cell cell_indx = row.getCell(1);
                list.add(new Ingredient(cell_name.getStringCellValue(), cell_indx.getStringCellValue()));
            }
        }
        return list;
    }

    public  ArrayList<Ingredient> getIdIngredient(ArrayList<Ingredient> arrayList) throws SQLException {

        ArrayList<Ingredient> list = new ArrayList<>();
        int minid = 359900000;
        int maxid = 359999990;
        String id1 = "3408";
        String id2 = "10";
        String id3 = "000";
        String id4 = "0";
        long id;
        String name;
        String index;
        String indexfin;
        String indexid;

        int i = 0;

        for (int a = 0; a < arrayList.size(); a++) {
            Ingredient ingredient = arrayList.get(a);
            index = ingredient.getIndx();

            indexfin = arrayList.get(a).getIndx();
            if(a > 0 && a < arrayList.size() - 1) {
                indexfin = arrayList.get(a + 1).getIndx();
            }
            name = ingredient.getName().trim();

            if (i < 99 && index.equals(indexfin)) {
                int j = i;
                if(i == 0) {
                    list.add(new Ingredient("***", "***")); //разделитель уровней
                    indexid = id1 + id2 + id3;
                    list.add(new Ingredient(indexid, index));
                    indexid = id1 + id2 + "0" + (j +1) + id4;
                    list.add(new Ingredient(indexid, name));
                } else if(i > 0 && i < 9){
                    indexid = id1 + id2 + "0" + (j +1) + id4;
                    list.add(new Ingredient(indexid, name));
                } else if(i >= 9) {
                    indexid = id1 + id2 + (j +1) + id4;
                    list.add(new Ingredient(indexid, name));
                }
                i++;
            } else if(i == 99 || !arrayList.get(a + 1).getIndx().equals(index)) {
                if(i > 0 && i < 9){
                    indexid = id1 + id2 + "0" + (i +1) + id4;
                    list.add(new Ingredient(indexid, arrayList.get(a).getName()));
                } else if(i >= 9) {
                    indexid = id1 + id2 + (i +1) + id4;
                    list.add(new Ingredient(indexid, arrayList.get(a).getName()));
                }
                int j = Integer.parseInt(id2) + 1;
                id2 = "" + j;
                i = 0;
            }
        }

        return list;
    }

    public static void main (String[] args) throws IOException, SQLException, ClassNotFoundException {

        IngInserter ingIns = new IngInserter();
        ArrayList<Ingredient> list = ingIns.getIngredients();
        int i = 0;
        int indx = 0;
        //for (Ingredient ingredient : list) {
        //    System.out.println(i);
        //    i++;
        //}
        FileInputStream inputStream = new FileInputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\Новая папка\\normative.xls");
        Workbook doc = new HSSFWorkbook(inputStream);
        Sheet sheetOut = doc.getSheet("2");

        for (Ingredient ingredient : ingIns.getIdIngredient(list)) {
            System.out.println(ingredient.getIndx() + " " + ingredient.getName());
            sheetOut.createRow(indx);
            Row rowOut = sheetOut.getRow(indx);
            rowOut.createCell(0);
            rowOut.createCell(1);
            rowOut.getCell(1).setCellValue(ingredient.getIndx());
            rowOut.getCell(0).setCellValue(ingredient.getName());
            indx++;
        }

        inputStream.close();
        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\Новая папка\\normative.xls");

        doc.write(fileOut);
        fileOut.close();
        doc.close();
        inputStream.close();
    }
}
