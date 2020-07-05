package padeg_exel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class Get_Person {

    ArrayList<Person> list_Person;

    Get_Person(ArrayList<Person> list_Person){
        this.list_Person = list_Person;
    }

    public ArrayList Get_Person() throws FileNotFoundException {

    final String FILENAME = "C:\\JAVA_EXEL\\Padeg_Exel\\22.xls";

    FileInputStream in = null;
    { try {in = new FileInputStream(FILENAME);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            System.out.println("Файл не найден");
        }
    }

    HSSFWorkbook wb = null;
    { try {wb = new HSSFWorkbook(in);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    Sheet sheet = wb.getSheetAt(0);

    //Переменные источника

    String name = "";
    double tabel  = 0;
    String S_Unit = "";
    String New_S_Unit = "";
    String Present_pos = "";
    String New_Present_pos = "";
    double prise = 0;
    String Prise = "";
    String date_new_transfer= null;
    double N_Ord = 0;
    String date_Ord = null;
    String date_OrdOrd = null;

    ArrayList<String> arrayList_NAME = new ArrayList<>();


    Row rowOrd = sheet.getRow(10);
    Cell DATE_ORDORD;
    if (rowOrd.getCell(34) == null || rowOrd.getCell(34).getCellType() == Cell.CELL_TYPE_BLANK){DATE_ORDORD = rowOrd.getCell(33);}
    else {DATE_ORDORD = rowOrd.getCell(34);}

    Iterator<Row> it = sheet.rowIterator();// Перебираем все строки
    if (it.hasNext()) {
        for (int i = 0; i < 18; i++) {
            it.next();
        }
        while (it.hasNext()) {
            Row row = it.next();
            Cell NAME = row.getCell(0);
            Cell TABEL = row.getCell(4);
            Cell S_UNIT = row.getCell(7);
            Cell New_S_UNIT = row.getCell(12);
            Cell PRESENT_POS = row.getCell(17);
            Cell NEW_PRESENT_POS = row.getCell(22);
            Cell PRISE = row.getCell(27);
            Cell DATE_NEW_TRANSFER = row.getCell(31);
            Cell N_ORD;
            if(row.getCell(38).getNumericCellValue() == 0.0){N_ORD = row.getCell(37);}
            else {N_ORD = row.getCell(38);}
            Cell DATE_ORD;
            if (row.getCell(42) == null || row.getCell(42).getCellType() == Cell.CELL_TYPE_BLANK){DATE_ORD = row.getCell(41);}
            else {DATE_ORD = row.getCell(42);}

            try {
                if (TABEL != null || TABEL.getCellType() != Cell.CELL_TYPE_BLANK) {
                    name = NAME.getStringCellValue();
                    //System.out.println(name);
                    if (TABEL.getCellTypeEnum() == CellType.NUMERIC) {tabel = TABEL.getNumericCellValue();}
                    else if (TABEL.getCellTypeEnum() == CellType.STRING) {tabel = Double.parseDouble(TABEL.getStringCellValue());}
                    //System.out.println(tabel);
                    S_Unit = S_UNIT.getStringCellValue();
                    New_S_Unit = New_S_UNIT.getStringCellValue();
                    Present_pos = PRESENT_POS.getStringCellValue();
                    New_Present_pos = NEW_PRESENT_POS.getStringCellValue();
                    if (PRISE.getCellTypeEnum() == CellType.NUMERIC) {prise = PRISE.getNumericCellValue();}
                    //else if (TABEL.getCellTypeEnum() == CellType.STRING) {prise = Double.parseDouble(TABEL.getStringCellValue());}
                    System.out.println(prise);
                    Prise = Integer.toString ((int)prise);
                    System.out.println(Prise);
                    StringBuffer sb = new StringBuffer(Prise);
                    if(sb.length() > 5){Prise = String.valueOf(sb.insert(3, " "));}
                    else Prise = String.valueOf(sb.insert(2, " "));
                    date_new_transfer = DATE_NEW_TRANSFER.getStringCellValue();
                    N_Ord = N_ORD.getNumericCellValue();
                    System.out.println(N_Ord);
                    date_Ord = DATE_ORD.getStringCellValue();
                    date_OrdOrd = DATE_ORDORD.getStringCellValue();
                    //System.out.println(name);
                    //arrayList_NAME.add(name);
                    list_Person.add(new Person(name, tabel, S_Unit, New_S_Unit, Present_pos, New_Present_pos, Prise, date_new_transfer, N_Ord, date_Ord, date_OrdOrd));
                } else {
                    break;
                }
            } catch (NullPointerException e) {
            }
        }
    }
    //for (int j = 0; j < arrayList_NAME.size(); j++){
        //System.out.println(arrayList_NAME.get(j));
        //list_Person.add(new Person(arrayList_NAME.get(j)));
    //}
        try {
            in.close();
            wb.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return list_Person;
    }
}
