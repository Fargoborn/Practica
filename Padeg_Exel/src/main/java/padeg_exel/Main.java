package padeg_exel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Iterator;


public class Main {

   public static void main  (String[] args) throws IOException, ParseException {


       String fio = "";
       int tabel  = 0;
       String S_Unit = "";
       String New_S_Unit = "";
       String Present_pos = "";
       String New_Present_pos = "";
       String Prise = "";
       String date_new_transfer= null;
       int N_Ord = 0;
       String date_Ord = null;
       String date_OrdOrd = null;
       int index_row = 3;

       Workbook workbook = new HSSFWorkbook();
       File file = new File("C:\\JAVA_EXEL\\Padeg_Exel\\Template_Ord.xls");
       FileOutputStream fileOutputStream = new FileOutputStream("C:\\JAVA_EXEL\\Padeg_Exel\\gen_Template_Ord.xls");
       Files.copy(file.toPath(), fileOutputStream);

       workbook.write(fileOutputStream);
       fileOutputStream.close();

       //Доступ к скопированному шаблону
       POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream("C:\\JAVA_EXEL\\Padeg_Exel\\gen_Template_Ord.xls"));
       Workbook workbook_templ = new HSSFWorkbook(fs, true);
       Sheet sheet = workbook_templ.getSheet("Лист1");

       Get_Person g_p = new Get_Person(new ArrayList<Person>());
       ArrayList list_Person = null;

       list_Person = new ArrayList(g_p.Get_Person());

       Iterator<Person> personIterator = list_Person.iterator();
       if (personIterator.hasNext()) {
           while (personIterator.hasNext()) {
               Person currPerson = personIterator.next();
               if (currPerson.get$name() != null){
                   Get_NativePadeg get_nativePadeg = new Get_NativePadeg(currPerson.get$name());
                   fio = get_nativePadeg.getGet_NativePadeg();
                   tabel = (int)currPerson.get$tabel();
                   S_Unit = currPerson.get$S_Unit();
                   New_S_Unit = currPerson.get$New_S_Unit();
                   Present_pos = (currPerson.get$Present_pos()).substring(0, currPerson.get$Present_pos().length() - 5);
                   New_Present_pos = (currPerson.get$New_Present_pos()).substring(0, currPerson.get$New_Present_pos().length() - 5);
                   Prise = currPerson.get$Prise();
                   date_new_transfer = currPerson.get$date_new_transfer();
                   N_Ord = (int)currPerson.get$N_Ord();
                   date_Ord = currPerson.$date_Ord;
                   date_OrdOrd = currPerson.$date_OrdOrd;

                   //System.out.println(fio + " " + date_new_transfer + " " + date_Ord);

                   //Запись в шаблон
                   Row row_FIO = sheet.getRow(index_row);
                   System.out.println(index_row + " "+ fio);
                   Cell cell_FIO = row_FIO.getCell(0);
                   cell_FIO.setCellValue(fio);

                   Row row_Tabel = sheet.getRow(index_row);
                   Cell cell_Tabel = row_Tabel.getCell(1);
                   cell_Tabel.setCellValue(tabel);

                   Row row_S_Unit = sheet.getRow(index_row);
                   Cell cell_S_Unit = row_S_Unit.getCell(2);
                   cell_S_Unit.setCellValue(S_Unit);

                   Row row_New_S_Unit = sheet.getRow(index_row);
                   Cell cell_New_S_Unit = row_New_S_Unit.getCell(3);
                   cell_New_S_Unit.setCellValue(New_S_Unit);

                   Row row_Present_pos = sheet.getRow(index_row);
                   Cell cell_Present_pos = row_Present_pos.getCell(4);
                   cell_Present_pos.setCellValue(Present_pos);

                   Row row_New_Present_pos = sheet.getRow(index_row);
                   Cell cell_New_Present_pos = row_New_Present_pos.getCell(5);
                   cell_New_Present_pos.setCellValue(New_Present_pos);

                   Row row_prise = sheet.getRow(index_row);
                   Cell cell_prise = row_prise.getCell(6);
                   cell_prise.setCellValue(Prise);

                   Row _row_date_new_transfer = sheet.getRow(index_row);
                   Cell _cell_date_new_transfer = _row_date_new_transfer.getCell(8);
                   _cell_date_new_transfer.setCellValue("---");

                   Row row_date_new_transfer = sheet.getRow(index_row);
                   Cell cell_date_new_transfer = row_date_new_transfer.getCell(7);
                   cell_date_new_transfer.setCellValue("постоянно с " + date_new_transfer);

                   Row row_N_Ord = sheet.getRow(index_row);
                   Cell cell_N_Ord = row_N_Ord.getCell(9);
                   cell_N_Ord.setCellValue("Доп.\n" + "соглашение к ТД №" + N_Ord + " от " + date_Ord);

                   Row row_date_Ord = sheet.getRow(index_row);
                   Cell cell_date_Ord = row_date_Ord.getCell(10);
                   cell_date_Ord.setCellValue(date_OrdOrd);

                   index_row++;
                   //sheet.createRow(index_row);
               }
           }
           try (FileOutputStream fileOut = new FileOutputStream("C:\\JAVA_EXEL\\Padeg_Exel\\gen_Template_Ord.xls")) {
               workbook_templ.write(fileOut);
               fileOut.close();
           }
           fs.close();
   }}}







