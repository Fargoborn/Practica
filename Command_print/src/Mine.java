
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.attribute.BasicFileAttributes;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;



public class Mine {

    public static void main  (String[] args) throws IOException, ParseException {

        InputStream in = null;
        HSSFWorkbook wb = null;


        final String FILENAME = "C:\\JAVA_EXEL\\ND.xls";

        Files.walk(Paths.get("C:\\JAVA_EXEL\\ORDER_EXEL"))
                .filter(Files::isRegularFile)
                .map(Path::toFile)
                .forEach(File::delete);
        try {
            in = new FileInputStream(FILENAME);
            try {
                wb = new HSSFWorkbook(in);

            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (FileNotFoundException ex) {
            System.out.print("Exception: no file");
        }


        Sheet sheet = wb.getSheetAt(0);

        double N_Ord = 0;
        Date date_Ord = null;
        double tabel  = 0;
        String name = "";
        String S_Unit = "";
        String Present_pos = "";
        String adres_city = "";
        String adres_in = "";
        String bisnes = "";
        Date date_s = null;
        Date date_f = null;


        SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy");
        SimpleDateFormat dateFormat_d = new SimpleDateFormat("dd");
        SimpleDateFormat dateFormat_m = new SimpleDateFormat("M");
        SimpleDateFormat dateFormat_y = new SimpleDateFormat("yy");

        String [] mon = {"января","февраля","марта","апреля","мая","июня","июля","августа","сентября","октября","ноября","декабря"};

        ArrayList<String> _fio = new ArrayList<>();
        ArrayList<String> _order = new ArrayList<>();

        ArrayList<Integer> arrayList_N_Ord = new ArrayList<>();
        ArrayList<Date> arrayList_DATE_Ord = new ArrayList<>();
        ArrayList<Integer> arrayList_TABEL = new ArrayList<>();
        ArrayList<String> arrayList_NAME = new ArrayList<>();
        ArrayList<String> arrayList_S_UNIT = new ArrayList<>();
        ArrayList<String> arrayList_PRESENT_POS = new ArrayList<>();
        ArrayList<String> arrayList_ADRES_CITY = new ArrayList<>();
        ArrayList<String> arrayList_ADRES_IN = new ArrayList<>();
        ArrayList<String> arrayList_BISNES = new ArrayList<>();
        ArrayList<Date> arrayList_DATE_S = new ArrayList<>();
        ArrayList<Date> arrayList_DATE_F = new ArrayList<>();

        ArrayList<Person> personArrayList = new ArrayList<>();
        ArrayList<Object> grup_personArrayList = new ArrayList<>();


        arrayList_N_Ord.add(0);// блокируем первую позицию массива Номер приказа
        arrayList_DATE_Ord.add(null);
        arrayList_TABEL.add(null);
        arrayList_NAME.add(null);
        arrayList_S_UNIT.add(null);
        arrayList_PRESENT_POS.add(null);
        arrayList_ADRES_CITY.add(null);
        arrayList_ADRES_IN.add(null);
        arrayList_BISNES.add(null);
        arrayList_DATE_S.add(null);
        arrayList_DATE_F.add(null);

        Iterator<Row> it = sheet.rowIterator();// Перебираем все строки

        if (it.hasNext()) {

            it.next();
            it.next();

            while (it.hasNext()) {
                Row row = it.next();


                Cell N_ORD = row.getCell(1);
                Cell DATE_Ord = row.getCell(0);
                Cell TABEL = row.getCell(2);
                Cell NAME = row.getCell(3);
                Cell S_UNIT = row.getCell(4);
                Cell PRESENT_POS = row.getCell(5);
                Cell ADRES_CITY = row.getCell(6);
                Cell ADRES_IN = row.getCell(7);
                Cell BISNES = row.getCell(8);
                Cell DATE_S = row.getCell(9);
                Cell DATE_F = row.getCell(10);

                //System.out.println(DATE_S.getColumnIndex() + " " + TABEL.getCellType());

                try{

                    if (N_ORD != null ){
                        if (N_ORD.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            N_Ord = N_ORD.getNumericCellValue();
                        } else {
                            N_Ord = Integer.parseInt(N_ORD.getStringCellValue().trim());
                        }
                        if (DATE_Ord.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            date_Ord = DATE_Ord.getDateCellValue();
                        } else {
                            SimpleDateFormat format = new SimpleDateFormat();
                            format.applyPattern("dd.MM.yyyy");
                            date_Ord = format.parse(DATE_Ord.getStringCellValue().trim());
                        }
                        if (TABEL.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            tabel = TABEL.getNumericCellValue();
                        } else {
                            tabel = Integer.parseInt(TABEL.getStringCellValue().trim());
                        }
                    name = NAME.getStringCellValue().trim();
                    S_Unit = S_UNIT.getStringCellValue().trim();
                    Present_pos = PRESENT_POS.getStringCellValue().trim();
                    adres_city = ADRES_CITY.getStringCellValue().trim();
                    adres_in = ADRES_IN.getStringCellValue().trim();
                    bisnes = BISNES.getStringCellValue();
                        if (DATE_S.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            date_s = DATE_S.getDateCellValue();
                        } else {
                            SimpleDateFormat format = new SimpleDateFormat();
                            format.applyPattern("dd.MM.yyyy");
                            date_s = format.parse(DATE_S.getStringCellValue().trim());
                        }
                        if (DATE_F.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            date_f = DATE_F.getDateCellValue();
                        } else {
                            SimpleDateFormat format = new SimpleDateFormat();
                            format.applyPattern("dd.MM.yyyy");
                            date_f = format.parse(DATE_F.getStringCellValue().trim());
                        }
                        System.out.println(N_Ord + " " + date_s + " " + date_f);
                    }
                    else {break;}
                }
                catch(NullPointerException e){}

                arrayList_N_Ord.add((int)N_Ord);
                arrayList_DATE_Ord.add(date_Ord);
                arrayList_TABEL.add((int)tabel);
                arrayList_NAME.add(name);
                arrayList_S_UNIT.add(S_Unit);
                arrayList_PRESENT_POS.add(Present_pos);
                arrayList_ADRES_CITY.add(adres_city);
                arrayList_ADRES_IN.add(adres_in);
                arrayList_BISNES.add(bisnes);
                arrayList_DATE_S.add(date_s);
                arrayList_DATE_F.add(date_f);
            }
        }

        arrayList_N_Ord.add(10000);// блокируем последнею позицию массива Номер приказа
        arrayList_DATE_Ord.add(null);
        arrayList_TABEL.add(null);
        arrayList_NAME.add(null);
        arrayList_S_UNIT.add(null);
        arrayList_PRESENT_POS.add(null);
        arrayList_ADRES_CITY.add(null);
        arrayList_ADRES_IN.add(null);
        arrayList_BISNES.add(null);
        arrayList_DATE_S.add(null);
        arrayList_DATE_F.add(null);

        for (int i = 1; i < arrayList_N_Ord.size() - 1; i++){

             if (!arrayList_N_Ord.get(i).equals(arrayList_N_Ord.get(i - 1)) && !arrayList_N_Ord.get(i).equals(arrayList_N_Ord.get(i + 1)))
            {
                int $N_Ord = (arrayList_N_Ord.get(i));
                String $date_Ord = dateFormat.format(arrayList_DATE_Ord.get(i));
                String $date_Ord_y = dateFormat_y.format(arrayList_DATE_Ord.get(i));

                int $tabel = (arrayList_TABEL.get(i));
                String $name = arrayList_NAME.get(i);
                String $S_Unit = (arrayList_S_UNIT.get(i));
                String $Present_pos = (arrayList_PRESENT_POS.get(i));
                String $adres_city = (arrayList_ADRES_CITY.get(i));
                String $adres_in = (arrayList_ADRES_IN.get(i));
                String $bisnes = (arrayList_BISNES.get(i));

                Date $date_S = arrayList_DATE_S.get(i);
                String $date_s_d = dateFormat_d.format(arrayList_DATE_S.get(i));
                String $date_s_m = dateFormat_m.format(arrayList_DATE_S.get(i));
                String $date_s_y = dateFormat_y.format(arrayList_DATE_S.get(i));

                Date $date_F = arrayList_DATE_F.get(i);
                String $date_f_d = dateFormat_d.format(arrayList_DATE_F.get(i));
                String $date_f_m = dateFormat_m.format(arrayList_DATE_F.get(i));
                String $date_f_y = dateFormat_y.format(arrayList_DATE_F.get(i));

                personArrayList.add(new Person($N_Ord, $date_Ord, $date_Ord_y, $tabel, $name, $S_Unit, $Present_pos, $adres_city, $adres_in, $bisnes, $date_S, $date_s_d, $date_s_m, $date_s_y, $date_F, $date_f_d, $date_f_m, $date_f_y));
            }

            else if (!arrayList_N_Ord.get(i).equals(arrayList_N_Ord.get(i - 1)) && arrayList_N_Ord.get(i).equals(arrayList_N_Ord.get(i + 1)))
            {
                int $1N_Ord = (arrayList_N_Ord.get(i));
                String $1date_Ord = dateFormat.format(arrayList_DATE_Ord.get(i));
                String $1date_Ord_y = dateFormat_y.format(arrayList_DATE_Ord.get(i));

                int $1tabel = (arrayList_TABEL.get(i));
                String $1name = arrayList_NAME.get(i);
                String $1S_Unit = (arrayList_S_UNIT.get(i));
                String $1Present_pos = (arrayList_PRESENT_POS.get(i));
                String $1adres_city = (arrayList_ADRES_CITY.get(i));
                String $1adres_in = (arrayList_ADRES_IN.get(i));
                String $1bisnes = (arrayList_BISNES.get(i));

                Date $1date_S = arrayList_DATE_S.get(i);
                String $1date_s_d = dateFormat_d.format(arrayList_DATE_S.get(i));
                String $1date_s_m = dateFormat_m.format(arrayList_DATE_S.get(i));
                String $1date_s_y = dateFormat_y.format(arrayList_DATE_S.get(i));

                Date $1date_F = arrayList_DATE_F.get(i);
                String $1date_f_d = dateFormat_d.format(arrayList_DATE_F.get(i));
                String $1date_f_m = dateFormat_m.format(arrayList_DATE_F.get(i));
                String $1date_f_y = dateFormat_y.format(arrayList_DATE_F.get(i));

                grup_personArrayList.add(new G1Person($1N_Ord, $1date_Ord, $1date_Ord_y, $1tabel, $1name, $1S_Unit, $1Present_pos, $1adres_city, $1adres_in, $1bisnes, $1date_S, $1date_s_d, $1date_s_m,
                        $1date_s_y, $1date_F, $1date_f_d, $1date_f_m, $1date_f_y));
            }

            else if ((arrayList_N_Ord.get(i).equals(arrayList_N_Ord.get(i - 1)) && arrayList_N_Ord.get(i).equals(arrayList_N_Ord.get(i + 1)))
                    || (arrayList_N_Ord.get(i).equals(arrayList_N_Ord.get(i - 1)) && !arrayList_N_Ord.get(i).equals(arrayList_N_Ord.get(i + 1)) && !arrayList_N_Ord.get(i).equals(arrayList_N_Ord.get(i - 2))))
            {
                int $2N_Ord = (arrayList_N_Ord.get(i));
                String $2date_Ord = dateFormat.format(arrayList_DATE_Ord.get(i));
                String $2date_Ord_y = dateFormat_y.format(arrayList_DATE_Ord.get(i));

                int $2tabel = (arrayList_TABEL.get(i));
                String $2name = arrayList_NAME.get(i);
                String $2S_Unit = (arrayList_S_UNIT.get(i));
                String $2Present_pos = (arrayList_PRESENT_POS.get(i));
                String $2adres_city = (arrayList_ADRES_CITY.get(i));
                String $2adres_in = (arrayList_ADRES_IN.get(i));
                String $2bisnes = (arrayList_BISNES.get(i));

                Date $2date_S = arrayList_DATE_S.get(i);
                String $2date_s_d = dateFormat_d.format(arrayList_DATE_S.get(i));
                String $2date_s_m = dateFormat_m.format(arrayList_DATE_S.get(i));
                String $2date_s_y = dateFormat_y.format(arrayList_DATE_S.get(i));

                Date $2date_F = arrayList_DATE_F.get(i);
                String $2date_f_d = dateFormat_d.format(arrayList_DATE_F.get(i));
                String $2date_f_m = dateFormat_m.format(arrayList_DATE_F.get(i));
                String $2date_f_y = dateFormat_y.format(arrayList_DATE_F.get(i));

                grup_personArrayList.add(new G2Person($2N_Ord, $2date_Ord, $2date_Ord_y, $2tabel, $2name, $2S_Unit, $2Present_pos, $2adres_city, $2adres_in, $2bisnes, $2date_S, $2date_s_d, $2date_s_m,
                        $2date_s_y, $2date_F, $2date_f_d, $2date_f_m, $2date_f_y));
            }

            else if (arrayList_N_Ord.get(i).equals(arrayList_N_Ord.get(i - 2)))
            {
                int $3N_Ord = (arrayList_N_Ord.get(i));
                String $3date_Ord = dateFormat.format(arrayList_DATE_Ord.get(i));
                String $3date_Ord_y = dateFormat_y.format(arrayList_DATE_Ord.get(i));

                int $3tabel = (arrayList_TABEL.get(i));
                String $3name = arrayList_NAME.get(i);
                String $3S_Unit = (arrayList_S_UNIT.get(i));
                String $3Present_pos = (arrayList_PRESENT_POS.get(i));
                String $3adres_city = (arrayList_ADRES_CITY.get(i));
                String $3adres_in = (arrayList_ADRES_IN.get(i));
                String $3bisnes = (arrayList_BISNES.get(i));

                Date $3date_S = arrayList_DATE_S.get(i);
                String $3date_s_d = dateFormat_d.format(arrayList_DATE_S.get(i));
                String $3date_s_m = dateFormat_m.format(arrayList_DATE_S.get(i));
                String $3date_s_y = dateFormat_y.format(arrayList_DATE_S.get(i));

                Date $3date_F = arrayList_DATE_F.get(i);
                String $3date_f_d = dateFormat_d.format(arrayList_DATE_F.get(i));
                String $3date_f_m = dateFormat_m.format(arrayList_DATE_F.get(i));
                String $3date_f_y = dateFormat_y.format(arrayList_DATE_F.get(i));

                grup_personArrayList.add(new G3Person($3N_Ord, $3date_Ord, $3date_Ord_y, $3tabel, $3name, $3S_Unit, $3Present_pos, $3adres_city, $3adres_in, $3bisnes, $3date_S, $3date_s_d, $3date_s_m,
                        $3date_s_y, $3date_F, $3date_f_d, $3date_f_m, $3date_f_y));
            }

            //вывод для проверки
            /*
            System.out.print($N_Ord + " ");
            System.out.print($date_Ord + " ");
            System.out.print($tabel + " ");
            System.out.print($name + " ");
            System.out.print($S_Unit + " ");
            System.out.print($Present_pos + " ");
            System.out.print($adres_in + " ");
            System.out.print($bisnes + " ");
            System.out.print($date_s_d + " " + $date_s_m + " " + $date_s_y + " ");
            System.out.print($date_f_d + " " + $date_f_m + " " + $date_f_y);
            */

            System.out.println("*Сортировка по номерам приказов*");
            in.close();
        }

        Iterator g_personIterator;// Перебираем все групповые Персоны
        g_personIterator = grup_personArrayList.iterator();

        if (g_personIterator.hasNext()) {
            while (g_personIterator.hasNext()) {
               Object g_currPerson = g_personIterator.next();
                System.out.println("Печать групповых приказов, шаблон t9a");
                if (g_currPerson instanceof G1Person) {
                    //System.out.println(((G1Person) g_currPerson).$1N_Ord + " 1группа" + g_currPerson);
                    //Копируем шаблон
                    URL url = Thread.currentThread().getContextClassLoader().getResource("C:\\JAVA_EXEL\\template_t9a.xls");
                    File file = new File("C:\\JAVA_EXEL\\template_t9a.xls");
                    FileOutputStream fileOutputStream = new FileOutputStream("C:\\JAVA_EXEL\\generate_t9a.xls");
                    Files.copy(file.toPath(), fileOutputStream);

                    Workbook workbook = new HSSFWorkbook();



                    workbook.write(fileOutputStream);
                    fileOutputStream.close();

                    //Запись общего массива (ФИО, Приказ)
                    _order.add(String.valueOf(((G1Person) g_currPerson).$1N_Ord));
                    _fio.add(((G1Person) g_currPerson).$1name);

                    //Доступ к скопированному шаблону
                    POIFSFileSystem fs = new POIFSFileSystem(
                            new FileInputStream("C:\\JAVA_EXEL\\generate_t9a.xls"));

                    Workbook workbook_templ = new HSSFWorkbook(fs, true);


                    Sheet sheet_templ = workbook_templ.getSheet("T-9a");
                    //*****
                    //запись номера приказа
                    int _$N_Ord = ((G1Person) g_currPerson).$1N_Ord;
                    Row row_N_Ord = sheet_templ.getRow(8);
                    Cell cell_N_Ord = row_N_Ord.getCell(60);
                    cell_N_Ord.setCellValue(_$N_Ord);

                    //запись даты приказа
                    String _$date_Ord = ((G1Person) g_currPerson).$1date_Ord;
                    Row row_date_Ord = sheet_templ.getRow(8);
                    Cell cell_date_Ord = row_date_Ord.getCell(79);
                    cell_date_Ord.setCellValue(_$date_Ord);

                    //запись табельного номера
                    int _$tabel = ((G1Person) g_currPerson).$1tabel;
                    Row row_tabel = sheet_templ.getRow(15);
                    Cell cell_tabel = row_tabel.getCell(43);
                    cell_tabel.setCellValue(_$tabel);

                    //запись ФИО
                    String _$name = ((G1Person) g_currPerson).$1name;
                    Row row_name = sheet_templ.getRow(14);
                    Cell cell_name = row_name.getCell(43);
                    cell_name.setCellValue(_$name);

                    //запись подразделения
                    String _$S_Unit = ((G1Person) g_currPerson).$1S_Unit;
                    Row row_S_Unit = sheet_templ.getRow(16);
                    Cell cell_S_Unit = row_S_Unit.getCell(43);
                    cell_S_Unit.setCellValue(_$S_Unit);

                    //запись должности
                    String _$Present_pos = ((G1Person) g_currPerson).$1Present_pos;
                    Row row_Present_pos = sheet_templ.getRow(17);
                    Cell cell_Present_pos = row_Present_pos.getCell(43);
                    cell_Present_pos.setCellValue(_$Present_pos);

                    //запись места назначения
                    String _$adres_city = ((G1Person) g_currPerson).$1adres_city;
                    Row row_adres_city = sheet_templ.getRow(18);
                    Cell cell_adres_city = row_adres_city.getCell(43);
                    cell_adres_city.setCellValue(_$adres_city);

                    String _$adres_in = ((G1Person) g_currPerson).$1adres_in;
                    Row row_adres_in = sheet_templ.getRow(19);
                    Cell cell_adres_in = row_adres_in.getCell(43);
                    cell_adres_in.setCellValue(_$adres_in);

                    //запись цели
                    String _$bisnes = ((G1Person) g_currPerson).$1bisnes;
                    Row row_bisnes = sheet_templ.getRow(23);
                    Cell cell_bisnes = row_bisnes.getCell(43);
                    cell_bisnes.setCellValue(_$bisnes);

                    //запись даты начало
                    Date _$date_s_d = ((G1Person) g_currPerson).$1date_S;
                    String ds = dateFormat.format(_$date_s_d);
                    Row row_date_s_d = sheet_templ.getRow(20);
                    Cell cell_date_s_d = row_date_s_d.getCell(43);
                    cell_date_s_d.setCellValue(ds);
                    /*

                    String _$date_s_d = ((G1Person) g_currPerson).$1date_s_d;
                    Row row_date_s_d = sheet_templ.getRow(32);
                    Cell cell_date_s_d = row_date_s_d.getCell(2);
                    cell_date_s_d.setCellValue(_$date_s_d);

                    String _$date_s_m = mon[Integer.parseInt(((G1Person) g_currPerson).$1date_s_m)];
                    Row row_date_s_m = sheet_templ.getRow(32);
                    Cell cell_date_s_m = row_date_s_m.getCell(5);
                    cell_date_s_m.setCellValue(_$date_s_m);

                    String _$date_s_y = ((G1Person) g_currPerson).$1date_s_y;
                    Row row_date_s_y = sheet_templ.getRow(32);
                    Cell cell_date_s_y = row_date_s_y.getCell(10);
                    cell_date_s_y.setCellValue(_$date_s_y);
                    */

                    //запись даты конец
                    Date _$date_f_d = ((G1Person) g_currPerson).$1date_F;
                    String df = dateFormat.format(_$date_f_d);
                    Row row_date_f_d = sheet_templ.getRow(21);
                    Cell cell_date_f_d = row_date_f_d.getCell(43);
                    cell_date_f_d.setCellValue(df);
                    /*
                    String _$date_f_d = ((G1Person) g_currPerson).$1date_f_d;
                    Row row_date_f_d = sheet_templ.getRow(32);
                    Cell cell_date_f_d = row_date_f_d.getCell(15);
                    cell_date_f_d.setCellValue(_$date_f_d);

                    Row row_date_f_m = sheet_templ.getRow(32);
                    Cell cell_date_f_m = row_date_f_m.getCell(18);
                    cell_date_f_m.setCellValue(mon[Integer.parseInt(((G1Person) g_currPerson).$1date_f_m)]);

                    Row row_date_f_y = sheet_templ.getRow(32);
                    Cell cell_date_f_y = row_date_f_y.getCell(23);
                    cell_date_f_y.setCellValue(((G1Person) g_currPerson).$1date_f_y);
                    */

                    //расчёт дней коммандировки
                    Date date1 = ((G1Person) g_currPerson).$1date_S;
                    Date date2 = ((G1Person) g_currPerson).$1date_F;

                    long milliseconds = date2.getTime() - date1.getTime();
                    // 24 часа = 1 440 минут = 1 день
                    int days = ((int) (milliseconds / (24 * 60 * 60 * 1000))) + 1;

                    Row row_days = sheet_templ.getRow(22);
                    Cell cell_days = row_days.getCell(43);
                    cell_days.setCellValue(days);

                    //Основание
                    Row row_base = sheet_templ.getRow(29);
                    Cell cell_base = row_base.getCell(35);
                    cell_base.setCellValue("служебное задание от" + " " + ((G1Person) g_currPerson).$1date_Ord);

                    FileOutputStream fileOut = new FileOutputStream("C:\\JAVA_EXEL\\ORDER_EXEL\\"  + ((G1Person) g_currPerson).$1N_Ord + "_" + ((G1Person) g_currPerson).$1date_Ord_y + ".xls");
                    workbook_templ.write(fileOut);
                    fileOut.close();
                    fs.close();
                }

                if (g_currPerson instanceof G2Person) {
                    //Доступ к скопированному шаблону
                    POIFSFileSystem fs = new POIFSFileSystem(
                            new FileInputStream("C:\\JAVA_EXEL\\ORDER_EXEL\\"  + ((G2Person) g_currPerson).$2N_Ord + "_" +((G2Person) g_currPerson).$2date_Ord_y + ".xls"));

                    Workbook workbook_templ = new HSSFWorkbook(fs, true);


                    Sheet sheet_templ = workbook_templ.getSheet("T-9a");

                    //Запись общего массива (ФИО, Приказ)
                    _order.add(String.valueOf(((G2Person) g_currPerson).$2N_Ord));
                    _fio.add(((G2Person) g_currPerson).$2name);

                    //запись номера приказа
                    int _$N_Ord = ((G2Person) g_currPerson).$2N_Ord;
                    Row row_N_Ord = sheet_templ.getRow(8);
                    Cell cell_N_Ord = row_N_Ord.getCell(60);
                    cell_N_Ord.setCellValue(_$N_Ord);

                    //запись даты приказа
                    String _$date_Ord = ((G2Person) g_currPerson).$2date_Ord;
                    Row row_date_Ord = sheet_templ.getRow(8);
                    Cell cell_date_Ord = row_date_Ord.getCell(79);
                    cell_date_Ord.setCellValue(_$date_Ord);

                    //запись табельного номера
                    int _$tabel = ((G2Person) g_currPerson).$2tabel;
                    Row row_tabel = sheet_templ.getRow(15);
                    Cell cell_tabel = row_tabel.getCell(65);
                    cell_tabel.setCellValue(_$tabel);

                    //запись ФИО
                    String _$name = ((G2Person) g_currPerson).$2name;
                    Row row_name = sheet_templ.getRow(14);
                    Cell cell_name = row_name.getCell(65);
                    cell_name.setCellValue(_$name);

                    //запись подразделения
                    String _$S_Unit = ((G2Person) g_currPerson).$2S_Unit;
                    Row row_S_Unit = sheet_templ.getRow(16);
                    Cell cell_S_Unit = row_S_Unit.getCell(65);
                    cell_S_Unit.setCellValue(_$S_Unit);

                    //запись должности
                    String _$Present_pos = ((G2Person) g_currPerson).$2Present_pos;
                    Row row_Present_pos = sheet_templ.getRow(17);
                    Cell cell_Present_pos = row_Present_pos.getCell(65);
                    cell_Present_pos.setCellValue(_$Present_pos);

                    //запись места назначения
                    String _$adres_city = ((G2Person) g_currPerson).$2adres_city;
                    Row row_adres_city = sheet_templ.getRow(18);
                    Cell cell_adres_city = row_adres_city.getCell(65);
                    cell_adres_city.setCellValue(_$adres_city);

                    String _$adres_in = ((G2Person) g_currPerson).$2adres_in;
                    Row row_adres_in = sheet_templ.getRow(19);
                    Cell cell_adres_in = row_adres_in.getCell(65);
                    cell_adres_in.setCellValue(_$adres_in);

                    //запись цели
                    String _$bisnes = ((G2Person) g_currPerson).$2bisnes;
                    Row row_bisnes = sheet_templ.getRow(23);
                    Cell cell_bisnes = row_bisnes.getCell(65);
                    cell_bisnes.setCellValue(_$bisnes);

                    //запись даты начало
                    Date _$date_s_d = ((G2Person) g_currPerson).$2date_S;
                    String ds = dateFormat.format(_$date_s_d);
                    Row row_date_s_d = sheet_templ.getRow(20);
                    Cell cell_date_s_d = row_date_s_d.getCell(65);
                    cell_date_s_d.setCellValue(ds);

                    /*
                    String _$date_s_d = ((G2Person) g_currPerson).$2date_s_d;
                    Row row_date_s_d = sheet_templ.getRow(32);
                    Cell cell_date_s_d = row_date_s_d.getCell(2);
                    cell_date_s_d.setCellValue(_$date_s_d);

                    String _$date_s_m = mon[Integer.parseInt(((G2Person) g_currPerson).$2date_s_m)];
                    Row row_date_s_m = sheet_templ.getRow(32);
                    Cell cell_date_s_m = row_date_s_m.getCell(5);
                    cell_date_s_m.setCellValue(_$date_s_m);

                    String _$date_s_y = ((G2Person) g_currPerson).$2date_s_y;
                    Row row_date_s_y = sheet_templ.getRow(32);
                    Cell cell_date_s_y = row_date_s_y.getCell(10);
                    cell_date_s_y.setCellValue(_$date_s_y);
                    */

                    //запись даты конец
                    Date _$date_f_d = ((G2Person) g_currPerson).$2date_F;
                    String df = dateFormat.format(_$date_f_d);
                    Row row_date_f_d = sheet_templ.getRow(21);
                    Cell cell_date_f_d = row_date_f_d.getCell(65);
                    cell_date_f_d.setCellValue(df);

                    /*
                    String _$date_f_d = ((G2Person) g_currPerson).$2date_f_d;
                    Row row_date_f_d = sheet_templ.getRow(32);
                    Cell cell_date_f_d = row_date_f_d.getCell(15);
                    cell_date_f_d.setCellValue(_$date_f_d);

                    Row row_date_f_m = sheet_templ.getRow(32);
                    Cell cell_date_f_m = row_date_f_m.getCell(18);
                    cell_date_f_m.setCellValue(mon[Integer.parseInt(((G2Person) g_currPerson).$2date_f_m)]);

                    Row row_date_f_y = sheet_templ.getRow(32);
                    Cell cell_date_f_y = row_date_f_y.getCell(23);
                    cell_date_f_y.setCellValue(((G2Person) g_currPerson).$2date_f_y);
                    */

                    //расчёт дней коммандировки
                    Date date1 = ((G2Person) g_currPerson).$2date_S;
                    Date date2 = ((G2Person) g_currPerson).$2date_F;

                    long milliseconds = date2.getTime() - date1.getTime();
                    // 24 часа = 1 440 минут = 1 день
                    int days = ((int) (milliseconds / (24 * 60 * 60 * 1000))) + 1;

                    Row row_days = sheet_templ.getRow(22);
                    Cell cell_days = row_days.getCell(65);
                    cell_days.setCellValue(days);

                    //Основание
                    Row row_base = sheet_templ.getRow(29);
                    Cell cell_base = row_base.getCell(35);
                    cell_base.setCellValue("служебное задание от" + " " + ((G2Person) g_currPerson).$2date_Ord);

                    FileOutputStream fileOut = new FileOutputStream("C:\\JAVA_EXEL\\ORDER_EXEL\\"  + ((G2Person) g_currPerson).$2N_Ord + "_" +((G2Person) g_currPerson).$2date_Ord_y + ".xls");
                    workbook_templ.write(fileOut);
                    fileOut.close();
                    fs.close();
                }

                if (g_currPerson instanceof G3Person) {
                    //Доступ к скопированному шаблону
                    POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream("C:\\JAVA_EXEL\\ORDER_EXEL\\"  + ((G3Person) g_currPerson).$3N_Ord + "_" +((G3Person) g_currPerson).$3date_Ord_y + ".xls"));

                    Workbook workbook_templ = new HSSFWorkbook(fs, true);


                    Sheet sheet_templ = workbook_templ.getSheet("T-9a");

                    //Запись общего массива (ФИО, Приказ)
                    _order.add(String.valueOf(((G3Person) g_currPerson).$3N_Ord));
                    _fio.add(((G3Person) g_currPerson).$3name);

                    //запись номера приказа
                    int _$N_Ord = ((G3Person) g_currPerson).$3N_Ord;
                    Row row_N_Ord = sheet_templ.getRow(8);
                    Cell cell_N_Ord = row_N_Ord.getCell(60);
                    cell_N_Ord.setCellValue(_$N_Ord);

                    //запись даты приказа
                    String _$date_Ord = ((G3Person) g_currPerson).$3date_Ord;
                    Row row_date_Ord = sheet_templ.getRow(8);
                    Cell cell_date_Ord = row_date_Ord.getCell(79);
                    cell_date_Ord.setCellValue(_$date_Ord);

                    //запись табельного номера
                    int _$tabel = ((G3Person) g_currPerson).$3tabel;
                    Row row_tabel = sheet_templ.getRow(15);
                    Cell cell_tabel = row_tabel.getCell(86);
                    cell_tabel.setCellValue(_$tabel);

                    //запись ФИО
                    String _$name = ((G3Person) g_currPerson).$3name;
                    Row row_name = sheet_templ.getRow(14);
                    Cell cell_name = row_name.getCell(86);
                    cell_name.setCellValue(_$name);

                    //запись подразделения
                    String _$S_Unit = ((G3Person) g_currPerson).$3S_Unit;
                    Row row_S_Unit = sheet_templ.getRow(16);
                    Cell cell_S_Unit = row_S_Unit.getCell(86);
                    cell_S_Unit.setCellValue(_$S_Unit);

                    //запись должности
                    String _$Present_pos = ((G3Person) g_currPerson).$3Present_pos;
                    Row row_Present_pos = sheet_templ.getRow(17);
                    Cell cell_Present_pos = row_Present_pos.getCell(86);
                    cell_Present_pos.setCellValue(_$Present_pos);

                    //запись места назначения
                    String _$adres_city = ((G3Person) g_currPerson).$3adres_city;
                    Row row_adres_city = sheet_templ.getRow(18);
                    Cell cell_adres_city = row_adres_city.getCell(86);
                    cell_adres_city.setCellValue(_$adres_city);

                    String _$adres_in = ((G3Person) g_currPerson).$3adres_in;
                    Row row_adres_in = sheet_templ.getRow(19);
                    Cell cell_adres_in = row_adres_in.getCell(86);
                    cell_adres_in.setCellValue(_$adres_in);

                    //запись цели
                    String _$bisnes = ((G3Person) g_currPerson).$3bisnes;
                    Row row_bisnes = sheet_templ.getRow(23);
                    Cell cell_bisnes = row_bisnes.getCell(86);
                    cell_bisnes.setCellValue(_$bisnes);

                    //запись даты начало
                    Date _$date_s_d = ((G3Person) g_currPerson).$3date_S;
                    String ds = dateFormat.format(_$date_s_d);
                    Row row_date_s_d = sheet_templ.getRow(20);
                    Cell cell_date_s_d = row_date_s_d.getCell(86);
                    cell_date_s_d.setCellValue(ds);


                    /*
                    String _$date_s_d = ((G3Person) g_currPerson).$3date_s_d;
                    Row row_date_s_d = sheet_templ.getRow(32);
                    Cell cell_date_s_d = row_date_s_d.getCell(2);
                    cell_date_s_d.setCellValue(_$date_s_d);

                    String _$date_s_m = mon[Integer.parseInt(((G3Person) g_currPerson).$3date_s_m)];
                    Row row_date_s_m = sheet_templ.getRow(32);
                    Cell cell_date_s_m = row_date_s_m.getCell(5);
                    cell_date_s_m.setCellValue(_$date_s_m);

                    String _$date_s_y = ((G3Person) g_currPerson).$3date_s_y;
                    Row row_date_s_y = sheet_templ.getRow(32);
                    Cell cell_date_s_y = row_date_s_y.getCell(10);
                    cell_date_s_y.setCellValue(_$date_s_y);
                    */

                    //запись даты конец
                    Date _$date_f_d = ((G3Person) g_currPerson).$3date_F;
                    String df = dateFormat.format(_$date_f_d);
                    Row row_date_f_d = sheet_templ.getRow(21);
                    Cell cell_date_f_d = row_date_f_d.getCell(86);
                    cell_date_f_d.setCellValue(df);


                    /*
                    String _$date_f_d = ((G3Person) g_currPerson).$3date_f_d;
                    Row row_date_f_d = sheet_templ.getRow(21);
                    Cell cell_date_f_d = row_date_f_d.getCell(86);
                    cell_date_f_d.setCellValue(_$date_f_d);

                    Row row_date_f_m = sheet_templ.getRow(21);
                    Cell cell_date_f_m = row_date_f_m.getCell(86);
                    cell_date_f_m.setCellValue(mon[Integer.parseInt(((G3Person) g_currPerson).$3date_f_m)]);

                    Row row_date_f_y = sheet_templ.getRow(32);
                    Cell cell_date_f_y = row_date_f_y.getCell(86);
                    cell_date_f_y.setCellValue(((G3Person) g_currPerson).$3date_f_y);
                    */

                    //расчёт дней коммандировки
                    Date date1 = ((G3Person) g_currPerson).$3date_S;
                    Date date2 = ((G3Person) g_currPerson).$3date_F;

                    long milliseconds = date2.getTime() - date1.getTime();
                    // 24 часа = 1 440 минут = 1 день
                    int days = ((int) (milliseconds / (24 * 60 * 60 * 1000))) + 1;

                    Row row_days = sheet_templ.getRow(22);
                    Cell cell_days = row_days.getCell(86);
                    cell_days.setCellValue(days);

                    //Основание
                    Row row_base = sheet_templ.getRow(29);
                    Cell cell_base = row_base.getCell(35);
                    cell_base.setCellValue("служебное задание от" + " " + ((G3Person) g_currPerson).$3date_Ord);

                    FileOutputStream fileOut = new FileOutputStream("C:\\JAVA_EXEL\\ORDER_EXEL\\"  + ((G3Person) g_currPerson).$3N_Ord + "_" +((G3Person) g_currPerson).$3date_Ord_y + ".xls");
                    workbook_templ.write(fileOut);
                    fileOut.close();
                    fs.close();
                }
            }
        }

            System.out.println("*_____________________________________*");
        Iterator<Person> personIterator;// Перебираем все одиночные Персоны
        personIterator = personArrayList.iterator();

        if (personIterator.hasNext()) {
            while (personIterator.hasNext()) {
                Person currPerson = personIterator.next();
                System.out.println("Печать одиночных приказов, шаблон t9");
            /*
            System.out.println(currPerson.$date_Ord);
            System.out.println(currPerson.$tabel);
            System.out.println(currPerson.$name);
            System.out.println(currPerson.$S_Unit);
            System.out.println(currPerson.$Present_pos);
            System.out.println(currPerson.$adres_in);
            System.out.println(currPerson.$bisnes);
            System.out.println(currPerson.$date_S);
            System.out.println(currPerson.$date_s_d);
            System.out.println(currPerson.$date_s_m);
            System.out.println(currPerson.$date_s_y);
            System.out.println(currPerson.$date_F);
            System.out.println(currPerson.$date_f_d);
            System.out.println(currPerson.$date_f_m);
            System.out.println(currPerson.$date_f_y);
            */

            //Копируем шаблон
            URL url1 = Thread.currentThread().getContextClassLoader().getResource("C:\\JAVA_EXEL\\template_t9.xls");
            File file1 = new File("C:\\JAVA_EXEL\\template_t9.xls");
            FileOutputStream fileOutputStream1 = new FileOutputStream("C:\\JAVA_EXEL\\generate.xls");
            Files.copy(file1.toPath(), fileOutputStream1);

            Workbook workbook1 = new HSSFWorkbook();
            workbook1.write(fileOutputStream1);
            fileOutputStream1.close();


            //Доступ к скопированному шаблону
            POIFSFileSystem fs1 = new POIFSFileSystem(
                    new FileInputStream("C:\\JAVA_EXEL\\generate.xls"));

            Workbook workbook_templ1 = new HSSFWorkbook(fs1, true);


            Sheet sheet_templ1 = workbook_templ1.getSheet("Т- 9");

            //Запись общего массива (ФИО, Приказ)
             _order.add(String.valueOf((currPerson).$N_Ord));
             _fio.add((currPerson).$name);

            //запись номера приказа
            int _$N_Ord = currPerson.$N_Ord;
            Row row_N_Ord = sheet_templ1.getRow(9);
            Cell cell_N_Ord = row_N_Ord.getCell(15);
            cell_N_Ord.setCellValue(_$N_Ord);

            //запись даты приказа
            String _$date_Ord = currPerson.$date_Ord;
            Row row_date_Ord = sheet_templ1.getRow(9);
            Cell cell_date_Ord = row_date_Ord.getCell(22);
            cell_date_Ord.setCellValue(_$date_Ord);

            //запись табельного номера
            int _$tabel = currPerson.$tabel;
            Row row_tabel = sheet_templ1.getRow(16);
            Cell cell_tabel = row_tabel.getCell(25);
            cell_tabel.setCellValue(_$tabel);

            //запись ФИО
            String _$name = currPerson.$name;
            Row row_name = sheet_templ1.getRow(17);
            Cell cell_name = row_name.getCell(0);
            cell_name.setCellValue(_$name);

            //запись подразделения
            String _$S_Unit = currPerson.$S_Unit;
            Row row_S_Unit = sheet_templ1.getRow(19);
            Cell cell_S_Unit = row_S_Unit.getCell(0);
            cell_S_Unit.setCellValue(_$S_Unit);

            //запись должности
            String _$Present_pos = currPerson.$Present_pos;
            Row row_Present_pos = sheet_templ1.getRow(21);
            Cell cell_Present_pos = row_Present_pos.getCell(0);
            cell_Present_pos.setCellValue(_$Present_pos);

            //запись места назначения
            String _$adres_in = currPerson.$adres_city + ", " + currPerson.$adres_in;
            Row row_adres_in = sheet_templ1.getRow(23);
            Cell cell_adres_in = row_adres_in.getCell(1);
            cell_adres_in.setCellValue(_$adres_in);

            //запись цели
            String _$bisnes = currPerson.$bisnes;
            Row row_bisnes = sheet_templ1.getRow(35);
            Cell cell_bisnes = row_bisnes.getCell(4);
            cell_bisnes.setCellValue(_$bisnes);

            //запись даты начало
            String _$date_s_d = currPerson.$date_s_d;
            Row row_date_s_d = sheet_templ1.getRow(32);
            Cell cell_date_s_d = row_date_s_d.getCell(2);
            cell_date_s_d.setCellValue(_$date_s_d);

            String _$date_s_m = mon[Integer.parseInt(currPerson.$date_s_m) - 1];
            Row row_date_s_m = sheet_templ1.getRow(32);
            Cell cell_date_s_m = row_date_s_m.getCell(5);
            cell_date_s_m.setCellValue(_$date_s_m);

            String _$date_s_y = currPerson.$date_s_y;
            Row row_date_s_y = sheet_templ1.getRow(32);
            Cell cell_date_s_y = row_date_s_y.getCell(10);
            cell_date_s_y.setCellValue(_$date_s_y);

            //запись даты конец
            String _$date_f_d = currPerson.$date_f_d;
            Row row_date_f_d = sheet_templ1.getRow(32);
            Cell cell_date_f_d = row_date_f_d.getCell(15);
            cell_date_f_d.setCellValue(_$date_f_d);

            Row row_date_f_m = sheet_templ1.getRow(32);
            Cell cell_date_f_m = row_date_f_m.getCell(18);
            cell_date_f_m.setCellValue(mon[Integer.parseInt(currPerson.$date_f_m) - 1]);

            Row row_date_f_y = sheet_templ1.getRow(32);
            Cell cell_date_f_y = row_date_f_y.getCell(23);
            cell_date_f_y.setCellValue(currPerson.$date_f_y);


            //расчёт дней коммандировки
            Date date1 = currPerson.$date_S;
            Date date2 = currPerson.$date_F;

            long milliseconds = date2.getTime() - date1.getTime();
            // 24 часа = 1 440 минут = 1 день
            int days = ((int) (milliseconds / (24 * 60 * 60 * 1000))) + 1;

            Row row_days = sheet_templ1.getRow(30);
            Cell cell_days = row_days.getCell(4);
            cell_days.setCellValue(days);

            //Основание

            Row row_base = sheet_templ1.getRow(41);
            Cell cell_base = row_base.getCell(8);
            cell_base.setCellValue("служебное задание от" + " " + currPerson.$date_Ord);


            //После записи в книге сохраняем, закрываем потоки
            FileOutputStream fileOut1 = new FileOutputStream("C:\\JAVA_EXEL\\ORDER_EXEL\\"  + currPerson.$N_Ord + "_" + currPerson.$date_Ord_y + ".xls");
            workbook_templ1.write(fileOut1);

            fileOut1.close();
            fs1.close();

            System.out.println("_____________________________________");

        }
      }

        //доп.приказы на выходные
        System.out.println("Печать приказов на выходные");

        //Печать выходных для одиночек

        String day_s = "будни";
        String day_f = "будни";

        Iterator personIterator_v;// Перебираем одиночек
        personIterator_v = personArrayList.iterator();
        if (personIterator_v.hasNext()) {
            while (personIterator_v.hasNext()) {
                Object currPerson = personIterator_v.next();

                day_s = "будни";
                day_f = "будни";

                    //Определение выходного начало
                    String __$date_s_y = ((Person) currPerson).$date_s_y;
                    String __$date_s_m = ((Person) currPerson).$date_s_m;
                    String __$date_s_d = ((Person) currPerson).$date_s_d;
                    String date = __$date_s_d + "." + __$date_s_m + "." + "20" + __$date_s_y;
                    // Переводим строку в дату
                    SimpleDateFormat format = new SimpleDateFormat("dd.MM.yyyy");
                    Date dayWeek = null;
                    dayWeek = format.parse(date);
                    day_s = new SimpleDateFormat("EEEE").format(dayWeek);

                    //System.out.println(((Person) currPerson).$N_Ord + "Person " + day_s);

                    //Определение выходного конец
                    String __$date_f_y = ((Person) currPerson).$date_f_y;
                    String __$date_f_m = ((Person) currPerson).$date_f_m;
                    String __$date_f_d = ((Person) currPerson).$date_f_d;
                    String date_fin = __$date_f_d + "." + __$date_f_m + "." + "20" + __$date_f_y;
                    // Переводим строку в дату
                    SimpleDateFormat format_f = new SimpleDateFormat("dd.MM.yyyy");
                    Date dayWeek_f = null;
                    dayWeek_f = format_f.parse(date_fin);
                    // Вывод дня недели даты на экран
                    day_f = new SimpleDateFormat("EEEE").format(dayWeek_f);

                    //System.out.println(((Person) currPerson).$N_Ord + "G1Person " + day_f);

                    //Копируем шаблон
                    //URL url_v = Thread.currentThread().getContextClassLoader().getResource("C:\\JAVA_EXEL\\template_v.xls");
                    File file_v = new File("C:\\JAVA_EXEL\\template_v.xls");
                    FileOutputStream fileOutputStream_v = new FileOutputStream("C:\\JAVA_EXEL\\generate_v.xls");
                    Files.copy(file_v.toPath(), fileOutputStream_v);

                    Workbook workbook_v = new HSSFWorkbook();


                    workbook_v.write(fileOutputStream_v);
                    fileOutputStream_v.close();

                    if (day_s.contains("вос") || day_s.contains("суб") || day_f.contains("вос") || day_f.contains("суб"))
                    {
                        //Доступ к скопированному шаблону
                        POIFSFileSystem fs_v = new POIFSFileSystem(new FileInputStream("C:\\JAVA_EXEL\\generate_v.xls"));
                        Workbook workbook_templ_v = new HSSFWorkbook(fs_v, true);

                        Sheet sheet_templ_v = workbook_templ_v.getSheet("v");

                        //Запись общего массива (ФИО, Приказ)
                        _order.add(String.valueOf(((Person) currPerson).$N_Ord) + "_вдк");
                        _fio.add(((Person) currPerson).$name);

                        //номер приказа
                        int _$N_Ord_v = ((Person) currPerson).$N_Ord;
                        Row row_N_Ord_v = sheet_templ_v.getRow(8);
                        Cell cell_N_Ord_v = row_N_Ord_v.getCell(10);
                        cell_N_Ord_v.setCellValue(_$N_Ord_v);

                        //дата приказа
                        String _$date_Ord_v = ((Person) currPerson).$date_Ord;
                        Row row_date_Ord_v = sheet_templ_v.getRow(8);
                        Cell cell_date_Ord_v = row_date_Ord_v.getCell(1);
                        cell_date_Ord_v.setCellValue(_$date_Ord_v);

                        //номер записи
                        String n = "1.";
                        Row n_v = sheet_templ_v.getRow(18);
                        Cell cell_n_v = n_v.getCell(0);
                        cell_n_v.setCellValue(n);

                        //основная запись
                        String line = "";
                        if ((day_s.contains("вос") || day_s.contains("суб")) && (day_f.contains("пон") || day_f.contains("вт")  || day_f.contains("ср")  || day_f.contains("чет")  || day_f.contains("пят")))
                        {
                            //System.out.println(day_s + " " + day_f);
                            line = ((Person) currPerson).$Present_pos + " " + ((Person) currPerson).$S_Unit + " " + ((Person) currPerson).$name +
                                    " " + "выехать в командировку в " + ((Person) currPerson).$adres_city + " в выходной день (" + day_s + ") " +
                                    ((Person) currPerson).$date_s_d + " " + mon[Integer.parseInt(((Person) currPerson).$date_s_m) - 1] + " 20" + ((Person) currPerson).$date_s_y + "г.";
                            Row line_v = sheet_templ_v.getRow(18);
                            Cell cell_line_v = line_v.getCell(1);
                            cell_line_v.setCellValue(line);
                        }

                        else if ((day_s.contains("вос") || day_s.contains("суб")) && (day_f.contains("вос") || day_f.contains("суб")))
                        {
                            //System.out.println(day_s + " " + day_f);
                            line = ((Person) currPerson).$Present_pos + " " + ((Person) currPerson).$S_Unit + " " + ((Person) currPerson).$name +
                                    " " + "выехать в командировку в " + ((Person) currPerson).$adres_city + " в выходной день (" + day_s + ") " +
                                    ((Person) currPerson).$date_s_d + " " + mon[Integer.parseInt(((Person) currPerson).$date_s_m) -1] + " 20" + ((Person) currPerson).$date_s_y + "г." +
                                    " и вернуться из командировки в " + "выходной день ("+ day_f + ") " + ((Person) currPerson).$date_f_d + " " +
                                    mon[Integer.parseInt(((Person) currPerson).$date_f_m) -1] + " 20" + ((Person) currPerson).$date_f_y + "г.";
                            Row line_v = sheet_templ_v.getRow(18);
                            Cell cell_line_v = line_v.getCell(1);
                            cell_line_v.setCellValue(line);
                        }
                        else if ((day_s.contains("пон") || day_s.contains("вт")  || day_s.contains("ср")  || day_s.contains("чет")  || day_s.contains("пят")) && (day_f.contains("вос") || day_f.contains("суб")))
                        {
                            //System.out.println(day_s + " " + day_f);
                            line = ((Person) currPerson).$Present_pos + " " + ((Person) currPerson).$S_Unit + " " + ((Person) currPerson).$name +
                                    " вернуться из командировки в " + ((Person) currPerson).$adres_city + ", в выходной день ("+ day_f + ") " + ((Person) currPerson).$date_f_d + " " +
                                    mon[Integer.parseInt(((Person) currPerson).$date_f_m) -1] + " 20" + ((Person) currPerson).$date_f_y + "г.";
                            Row line_v = sheet_templ_v.getRow(18);
                            Cell cell_line_v = line_v.getCell(1);
                            cell_line_v.setCellValue(line);
                        }

                        //с приказом ознакомлен
                        String end = "С приказом ознакомлен, согласен: ____________  " + ((Person) currPerson).$name;
                        Row end_v = sheet_templ_v.getRow(26);
                        Cell cell_end_v = end_v.getCell(0);
                        cell_end_v.setCellValue(end);

                        //для бухгалтерии
                        String nn = "2.";
                        Row nn_v = sheet_templ_v.getRow(21);
                        Cell cell_nn_v = nn_v.getCell(0);
                        cell_nn_v.setCellValue(nn);


                        FileOutputStream fileOut = new FileOutputStream("C:\\JAVA_EXEL\\ORDER_EXEL\\"  + ((Person) currPerson).$N_Ord + "_вдк_" +((Person) currPerson).$date_Ord_y + ".xls");
                        workbook_templ_v.write(fileOut);
                        fileOut.close();
                        fs_v.close();

                    }

                }
        }


        //Печать выходных для групповых приказов
        Iterator g_personIterator_v;// Перебираем все групповые Персоны
        g_personIterator_v = grup_personArrayList.iterator();

        if (g_personIterator_v.hasNext()) {
            while (g_personIterator_v.hasNext()) {
                Object g_currPerson = g_personIterator_v.next();


                if (g_currPerson instanceof G1Person) {

                    day_s = "будни";
                    day_f = "будни";


                    //Определение выходного начало
                    String __$date_s_y = ((G1Person) g_currPerson).$1date_s_y;
                    String __$date_s_m = ((G1Person) g_currPerson).$1date_s_m;
                    String __$date_s_d = ((G1Person) g_currPerson).$1date_s_d;
                    String date = __$date_s_d + "." + __$date_s_m + "." + "20" + __$date_s_y;
                    // Переводим строку в дату
                    SimpleDateFormat format = new SimpleDateFormat("dd.MM.yyyy");
                    Date dayWeek = null;
                    dayWeek = format.parse(date);
                    day_s = new SimpleDateFormat("EEEE").format(dayWeek);

                    //System.out.println(((G1Person) g_currPerson).$1N_Ord + "G1Person " + day_s);

                    //Определение выходного конец
                    String __$date_f_y = ((G1Person) g_currPerson).$1date_f_y;
                    String __$date_f_m = ((G1Person) g_currPerson).$1date_f_m;
                    String __$date_f_d = ((G1Person) g_currPerson).$1date_f_d;
                    String date_fin = __$date_f_d + "." + __$date_f_m + "." + "20" + __$date_f_y;
                    // Переводим строку в дату
                    SimpleDateFormat format_f = new SimpleDateFormat("dd.MM.yyyy");
                    Date dayWeek_f = null;
                    dayWeek_f = format_f.parse(date_fin);
                    // Вывод дня недели даты на экран
                    day_f = new SimpleDateFormat("EEEE").format(dayWeek_f);

                    //System.out.println(((G1Person) g_currPerson).$1N_Ord + "G1Person " + day_f);

                    //Копируем шаблон
                    //URL url_v = Thread.currentThread().getContextClassLoader().getResource("C:\\JAVA_EXEL\\template_v.xls");
                    File file_v = new File("C:\\JAVA_EXEL\\template_v.xls");
                    FileOutputStream fileOutputStream_v = new FileOutputStream("C:\\JAVA_EXEL\\generate_v.xls");
                    Files.copy(file_v.toPath(), fileOutputStream_v);

                    Workbook workbook_v = new HSSFWorkbook();
                    workbook_v.write(fileOutputStream_v);
                    fileOutputStream_v.close();

                    if (day_s.contains("вос") || day_s.contains("суб") || day_f.contains("вос") || day_f.contains("суб"))
                    {
                        //Доступ к скопированному шаблону
                        POIFSFileSystem fs_v = new POIFSFileSystem(new FileInputStream("C:\\JAVA_EXEL\\generate_v.xls"));
                        Workbook workbook_templ_v = new HSSFWorkbook(fs_v, true);


                        Sheet sheet_templ_v = workbook_templ_v.getSheet("v");

                        //Запись общего массива (ФИО, Приказ)
                        _order.add(String.valueOf(((G1Person) g_currPerson).$1N_Ord) + "_вдк");
                        _fio.add(((G1Person) g_currPerson).$1name);

                        //номер приказа
                        int _$N_Ord_v = ((G1Person) g_currPerson).$1N_Ord;
                        Row row_N_Ord_v = sheet_templ_v.getRow(8);
                        Cell cell_N_Ord_v = row_N_Ord_v.getCell(10);
                        cell_N_Ord_v.setCellValue(_$N_Ord_v);

                        //дата приказа
                        String _$date_Ord_v = ((G1Person) g_currPerson).$1date_Ord;
                        Row row_date_Ord_v = sheet_templ_v.getRow(8);
                        Cell cell_date_Ord_v = row_date_Ord_v.getCell(1);
                        cell_date_Ord_v.setCellValue(_$date_Ord_v);

                        //номер записи
                        String n = "1.";
                        Row n_v = sheet_templ_v.getRow(18);
                        Cell cell_n_v = n_v.getCell(0);
                        cell_n_v.setCellValue(n);

                        //основная запись
                        String line = "";
                        if ((day_s.contains("вос") || day_s.contains("суб")) && (day_f.contains("пон") || day_f.contains("вт")  || day_f.contains("ср")  || day_f.contains("чет")  || day_f.contains("пят")))
                        {
                            //System.out.println(day_s + " " + day_f);
                            line = ((G1Person) g_currPerson).$1Present_pos + " " + ((G1Person) g_currPerson).$1S_Unit + " " + ((G1Person) g_currPerson).$1name +
                                " " + "выехать в командировку в " + ((G1Person) g_currPerson).$1adres_city + " в выходной день (" + day_s + ") " +
                                        ((G1Person) g_currPerson).$1date_s_d + " " + mon[Integer.parseInt(((G1Person) g_currPerson).$1date_s_m) - 1] + " 20" + ((G1Person) g_currPerson).$1date_s_y + "г.";
                            Row line_v = sheet_templ_v.getRow(18);
                            Cell cell_line_v = line_v.getCell(1);
                            cell_line_v.setCellValue(line);
                        }

                        else if ((day_s.contains("вос") || day_s.contains("суб")) && (day_f.contains("вос") || day_f.contains("суб")))
                         {
                             //System.out.println(day_s + " " + day_f);
                             line = ((G1Person) g_currPerson).$1Present_pos + " " + ((G1Person) g_currPerson).$1S_Unit + " " + ((G1Person) g_currPerson).$1name +
                                    " " + "выехать в командировку в " + ((G1Person) g_currPerson).$1adres_city + " в выходной день (" + day_s + ") " +
                                    ((G1Person) g_currPerson).$1date_s_d + " " + mon[Integer.parseInt(((G1Person) g_currPerson).$1date_s_m) -1] + " 20" + ((G1Person) g_currPerson).$1date_s_y + "г." +
                                    " и вернуться из командировки в " + "выходной день ("+ day_f + ") " + ((G1Person) g_currPerson).$1date_f_d + " " +
                                    mon[Integer.parseInt(((G1Person) g_currPerson).$1date_f_m) -1] + " 20" + ((G1Person) g_currPerson).$1date_f_y + "г.";
                            Row line_v = sheet_templ_v.getRow(18);
                            Cell cell_line_v = line_v.getCell(1);
                            cell_line_v.setCellValue(line);
                        }
                        else if ((day_s.contains("пон") || day_s.contains("вт")  || day_s.contains("ср")  || day_s.contains("чет")  || day_s.contains("пят")) && (day_f.contains("вос") || day_f.contains("суб")))
                         {
                             //System.out.println(day_s + " " + day_f);
                             line = ((G1Person) g_currPerson).$1Present_pos + " " + ((G1Person) g_currPerson).$1S_Unit + " " + ((G1Person) g_currPerson).$1name +
                                    " вернуться из командировки в " + ((G1Person) g_currPerson).$1adres_city + ", в выходной день ("+ day_f + ") " + ((G1Person) g_currPerson).$1date_f_d + " " +
                                    mon[Integer.parseInt(((G1Person) g_currPerson).$1date_f_m) -1] + " 20" + ((G1Person) g_currPerson).$1date_f_y + "г.";
                            Row line_v = sheet_templ_v.getRow(18);
                            Cell cell_line_v = line_v.getCell(1);
                            cell_line_v.setCellValue(line);
                        }

                        //с приказом ознакомлен
                        String end = "С приказом ознакомлен, согласен: ____________  " + ((G1Person) g_currPerson).$1name;
                        Row end_v = sheet_templ_v.getRow(26);
                        Cell cell_end_v = end_v.getCell(0);
                        cell_end_v.setCellValue(end);

                        //для бухгалтерии
                        String nn = "2.";
                        Row nn_v = sheet_templ_v.getRow(21);
                        Cell cell_nn_v = nn_v.getCell(0);
                        cell_nn_v.setCellValue(nn);


                        FileOutputStream fileOut = new FileOutputStream("C:\\JAVA_EXEL\\ORDER_EXEL\\"  + ((G1Person) g_currPerson).$1N_Ord + "_вдк_" +((G1Person) g_currPerson).$1date_Ord_y + ".xls");
                        workbook_templ_v.write(fileOut);
                        fileOut.close();
                        fs_v.close();

                    }

                }

                if (g_currPerson instanceof G2Person) {

                    day_s = "будни";
                    day_f = "будни";

                    //Определение выходного начало
                    String __$date_s_y = ((G2Person) g_currPerson).$2date_s_y;
                    String __$date_s_m = ((G2Person) g_currPerson).$2date_s_m;
                    String __$date_s_d = ((G2Person) g_currPerson).$2date_s_d;
                    String date = __$date_s_d + "." + __$date_s_m + "." + "20" + __$date_s_y;
                    // Переводим строку в дату
                    SimpleDateFormat format = new SimpleDateFormat("dd.MM.yyyy");
                    Date dayWeek = null;
                    dayWeek = format.parse(date);
                    // Вывод дня недели даты на экран
                    day_s = new SimpleDateFormat("EEEE").format(dayWeek);

                    //System.out.println(((G2Person) g_currPerson).$2N_Ord + "G2Person " + day_s);

                    //Определение выходного конец
                    String __$date_f_y = ((G2Person) g_currPerson).$2date_f_y;
                    String __$date_f_m = ((G2Person) g_currPerson).$2date_f_m;
                    String __$date_f_d = ((G2Person) g_currPerson).$2date_f_d;
                    String date_fin = __$date_f_d + "." + __$date_f_m + "." + "20" + __$date_f_y;
                    // Переводим строку в дату
                    SimpleDateFormat format_f = new SimpleDateFormat("dd.MM.yyyy");
                    Date dayWeek_f = null;
                    dayWeek_f = format_f.parse(date_fin);
                    // Вывод дня недели даты на экран
                    day_f = new SimpleDateFormat("EEEE").format(dayWeek_f);

                    //System.out.println(((G2Person) g_currPerson).$2N_Ord + "G2Person " + day_f);

                    File file = new File("C:\\JAVA_EXEL\\ORDER_EXEL\\"  + ((G2Person) g_currPerson).$2N_Ord + "_вдк_" +((G2Person) g_currPerson).$2date_Ord_y + ".xls");
                    if (file.exists()) {
                        POIFSFileSystem fs_v = new POIFSFileSystem(new FileInputStream("C:\\JAVA_EXEL\\ORDER_EXEL\\"  + ((G2Person) g_currPerson).$2N_Ord + "_вдк_" +((G2Person) g_currPerson).$2date_Ord_y + ".xls"));
                        Workbook workbook_templ_v = new HSSFWorkbook(fs_v, true);


                        Sheet sheet_templ_v = workbook_templ_v.getSheet("v");

                        if (day_s.contains("вос") || day_s.contains("суб") || day_f.contains("вос") || day_f.contains("суб")){
                            //Доступ к скопированному шаблону

                            //Запись общего массива (ФИО, Приказ)

                            _order.add(String.valueOf(((G2Person) g_currPerson).$2N_Ord) + "_вдк");
                            _fio.add(((G2Person) g_currPerson).$2name);

                            //номер приказа
                            int _$N_Ord_v = ((G2Person) g_currPerson).$2N_Ord;
                            Row row_N_Ord_v = sheet_templ_v.getRow(8);
                            Cell cell_N_Ord_v = row_N_Ord_v.getCell(10);
                            cell_N_Ord_v.setCellValue(_$N_Ord_v);

                            //дата приказа
                            String _$date_Ord_v = ((G2Person) g_currPerson).$2date_Ord;
                            Row row_date_Ord_v = sheet_templ_v.getRow(8);
                            Cell cell_date_Ord_v = row_date_Ord_v.getCell(1);
                            cell_date_Ord_v.setCellValue(_$date_Ord_v);

                            //номер записи
                            Row _n_v = sheet_templ_v.getRow(18);
                            Cell _cell_n_v = _n_v.getCell(0);
                            //System.out.println(_cell_n_v.getColumnIndex() + " " + _cell_n_v.getCellType());
                            String n_n_v = _cell_n_v.getStringCellValue();

                            if(n_n_v.equals("1."))  {
                                String n = "2.";
                                Row n_v = sheet_templ_v.getRow(19);
                                Cell cell_n_v = n_v.getCell(0);
                                cell_n_v.setCellValue(n);

                                //основная запись
                                String line = "";
                                if ((day_s.contains("вос") || day_s.contains("суб")) && (day_f.contains("пон") || day_f.contains("вт")  || day_f.contains("ср")  || day_f.contains("чет")  || day_f.contains("пят"))){
                                    line = ((G2Person) g_currPerson).$2Present_pos + " " + ((G2Person) g_currPerson).$2S_Unit + " " + ((G2Person) g_currPerson).$2name +
                                            " " + "выехать в командировку в " + ((G2Person) g_currPerson).$2adres_city + " в выходной день (" + day_s + ") " +
                                            ((G2Person) g_currPerson).$2date_s_d + " " + mon[Integer.parseInt(((G2Person) g_currPerson).$2date_s_m) - 1] + " 20" + ((G2Person) g_currPerson).$2date_s_y + "г.";
                                    Row line_v = sheet_templ_v.getRow(19);
                                    Cell cell_line_v = line_v.getCell(1);
                                    cell_line_v.setCellValue(line);}

                                else if ((day_s.contains("вос") || day_s.contains("суб")) && (day_f.contains("вос") || day_f.contains("суб"))){
                                    line = ((G2Person) g_currPerson).$2Present_pos + " " + ((G2Person) g_currPerson).$2S_Unit + " " + ((G2Person) g_currPerson).$2name +
                                            " " + "выехать в командировку в " + ((G2Person) g_currPerson).$2adres_city + " в выходной день (" + day_s + ") " +
                                            ((G2Person) g_currPerson).$2date_s_d + " " + mon[Integer.parseInt(((G2Person) g_currPerson).$2date_s_m) -1] + " 20" + ((G2Person) g_currPerson).$2date_s_y + "г." +
                                            " и вернуться из командировки в " + "выходной день ("+ day_f + ") " + ((G2Person) g_currPerson).$2date_f_d + " " +
                                            mon[Integer.parseInt(((G2Person) g_currPerson).$2date_f_m) -1] + " 20" + ((G2Person) g_currPerson).$2date_f_y + "г.";
                                    Row line_v = sheet_templ_v.getRow(19);
                                    Cell cell_line_v = line_v.getCell(1);
                                    cell_line_v.setCellValue(line);
                                }
                                else if ((day_s.contains("пон") || day_s.contains("вт")  || day_s.contains("ср")  || day_s.contains("чет")  || day_s.contains("пят")) && (day_f.contains("вос") || day_f.contains("суб"))){
                                    line = ((G2Person) g_currPerson).$2Present_pos + " " + ((G2Person) g_currPerson).$2S_Unit + " " + ((G2Person) g_currPerson).$2name +
                                            " вернуться из командировки в " + ((G2Person) g_currPerson).$2adres_city + ", в выходной день ("+ day_f + ") " + ((G2Person) g_currPerson).$2date_f_d + " " +
                                            mon[Integer.parseInt(((G2Person) g_currPerson).$2date_f_m) -1] + " 20" + ((G2Person) g_currPerson).$2date_f_y + "г.";
                                    Row line_v = sheet_templ_v.getRow(19);
                                    Cell cell_line_v = line_v.getCell(1);
                                    cell_line_v.setCellValue(line);
                                }

                                //с приказом ознакомлен
                                String end = "С приказом ознакомлен, согласен: ____________  " + ((G2Person) g_currPerson).$2name;
                                Row end_v = sheet_templ_v.getRow(27);
                                Cell cell_end_v = end_v.getCell(0);
                                cell_end_v.setCellValue(end);

                                //для бухгалтерии
                                String nn = "3.";
                                Row nn_v = sheet_templ_v.getRow(21);
                                Cell cell_nn_v = nn_v.getCell(0);
                                cell_nn_v.setCellValue(nn);

                            }

                            FileOutputStream fileOut = new FileOutputStream("C:\\JAVA_EXEL\\ORDER_EXEL\\"  + ((G2Person) g_currPerson).$2N_Ord + "_вдк_" +((G2Person) g_currPerson).$2date_Ord_y + ".xls");
                            workbook_templ_v.write(fileOut);
                            fileOut.close();
                            fs_v.close();

                        }

                    }

                    else {
                        POIFSFileSystem fs_v = new POIFSFileSystem(new FileInputStream("C:\\JAVA_EXEL\\generate_v.xls"));
                        Workbook workbook_templ_v = new HSSFWorkbook(fs_v, true);

                        Sheet sheet_templ_v = workbook_templ_v.getSheet("v");

                        if (day_s.contains("вос") || day_s.contains("суб") || day_f.contains("вос") || day_f.contains("суб")){
                        //Доступ к скопированному шаблону

                        //номер приказа
                        int _$N_Ord_v = ((G2Person) g_currPerson).$2N_Ord;
                        Row row_N_Ord_v = sheet_templ_v.getRow(8);
                        Cell cell_N_Ord_v = row_N_Ord_v.getCell(10);
                        cell_N_Ord_v.setCellValue(_$N_Ord_v);

                        //дата приказа
                        String _$date_Ord_v = ((G2Person) g_currPerson).$2date_Ord;
                        Row row_date_Ord_v = sheet_templ_v.getRow(8);
                        Cell cell_date_Ord_v = row_date_Ord_v.getCell(1);
                        cell_date_Ord_v.setCellValue(_$date_Ord_v);

                        //номер записи
                            String n = "1.";
                            Row n_v = sheet_templ_v.getRow(18);
                            Cell cell_n_v = n_v.getCell(0);
                            cell_n_v.setCellValue(n);

                            //основная запись
                            String line = "";
                            if ((day_s.contains("вос") || day_s.contains("суб")) && (day_f.contains("пон") || day_f.contains("вт")  || day_f.contains("ср")  || day_f.contains("чет")  || day_f.contains("пят"))){
                                line = ((G2Person) g_currPerson).$2Present_pos + " " + ((G2Person) g_currPerson).$2S_Unit + " " + ((G2Person) g_currPerson).$2name +
                                        " " + "выехать в командировку в " + ((G2Person) g_currPerson).$2adres_city + " в выходной день (" + day_s + ") " +
                                        ((G2Person) g_currPerson).$2date_s_d + " " + mon[Integer.parseInt(((G2Person) g_currPerson).$2date_s_m) - 1] + " 20" + ((G2Person) g_currPerson).$2date_s_y + "г.";
                                Row line_v = sheet_templ_v.getRow(18);
                                Cell cell_line_v = line_v.getCell(1);
                                cell_line_v.setCellValue(line);}

                            else if ((day_s.contains("вос") || day_s.contains("суб")) && (day_f.contains("вос") || day_f.contains("суб"))){
                                line = ((G2Person) g_currPerson).$2Present_pos + " " + ((G2Person) g_currPerson).$2S_Unit + " " + ((G2Person) g_currPerson).$2name +
                                        " " + "выехать в командировку в " + ((G2Person) g_currPerson).$2adres_city + " в выходной день (" + day_s + ") " +
                                        ((G2Person) g_currPerson).$2date_s_d + " " + mon[Integer.parseInt(((G2Person) g_currPerson).$2date_s_m) -1] + " 20" + ((G2Person) g_currPerson).$2date_s_y + "г." +
                                        " и вернуться из командировки в " + "в выходной день ("+ day_f + ") " + ((G2Person) g_currPerson).$2date_f_d + " " +
                                        mon[Integer.parseInt(((G2Person) g_currPerson).$2date_f_m) -1] + " 20" + ((G2Person) g_currPerson).$2date_f_y + "г.";
                                Row line_v = sheet_templ_v.getRow(18);
                                Cell cell_line_v = line_v.getCell(1);
                                cell_line_v.setCellValue(line);
                            }
                            else if ((day_s.contains("пон") || day_s.contains("вт")  || day_s.contains("ср")  || day_s.contains("чет")  || day_s.contains("пят")) && (day_f.contains("вос") || day_f.contains("суб"))){
                                line = ((G2Person) g_currPerson).$2Present_pos + " " + ((G2Person) g_currPerson).$2S_Unit + " " + ((G2Person) g_currPerson).$2name +
                                        " вернуться из командировки в " + ((G2Person) g_currPerson).$2adres_city + ", в выходной день ("+ day_f + ") " + ((G2Person) g_currPerson).$2date_f_d + " " +
                                        mon[Integer.parseInt(((G2Person) g_currPerson).$2date_f_m) -1] + " 20" + ((G2Person) g_currPerson).$2date_f_y + "г.";
                                Row line_v = sheet_templ_v.getRow(18);
                                Cell cell_line_v = line_v.getCell(1);
                                cell_line_v.setCellValue(line);
                            }

                            //с приказом ознакомлен
                            String end = "С приказом ознакомлен, согласен: ____________  " + ((G2Person) g_currPerson).$2name;
                            Row end_v = sheet_templ_v.getRow(26);
                            Cell cell_end_v = end_v.getCell(0);
                            cell_end_v.setCellValue(end);

                            //для бухгалтерии
                            String nn = "2.";
                            Row nn_v = sheet_templ_v.getRow(21);
                            Cell cell_nn_v = nn_v.getCell(0);
                            cell_nn_v.setCellValue(nn);

                        FileOutputStream fileOut = new FileOutputStream("C:\\JAVA_EXEL\\ORDER_EXEL\\"  + ((G2Person) g_currPerson).$2N_Ord + "_вдк_" +((G2Person) g_currPerson).$2date_Ord_y + ".xls");
                        workbook_templ_v.write(fileOut);
                        fileOut.close();
                        fs_v.close();

                    }

                  }
                }
                    if (g_currPerson instanceof G3Person) {

                        day_s = "будни";
                        day_f = "будни";

                        //Определение выходного начало
                        String __$date_s_y = ((G3Person) g_currPerson).$3date_s_y;
                        String __$date_s_m = ((G3Person) g_currPerson).$3date_s_m;
                        String __$date_s_d = ((G3Person) g_currPerson).$3date_s_d;
                        String date = __$date_s_d + "." + __$date_s_m + "." + "20" + __$date_s_y;
                        // Переводим строку в дату
                        SimpleDateFormat format = new SimpleDateFormat("dd.MM.yyyy");
                        Date dayWeek = null;
                        dayWeek = format.parse(date);
                        // Вывод дня недели даты на экран
                        day_s = new SimpleDateFormat("EEEE").format(dayWeek);

                        //System.out.println(((G3Person) g_currPerson).$3N_Ord + "G3Person " + day_s);

                        //Определение выходного конец
                        String __$date_f_y = ((G3Person) g_currPerson).$3date_f_y;
                        String __$date_f_m = ((G3Person) g_currPerson).$3date_f_m;
                        String __$date_f_d = ((G3Person) g_currPerson).$3date_f_d;
                        String date_fin = __$date_f_d + "." + __$date_f_m + "." + "20" + __$date_f_y;
                        // Переводим строку в дату
                        SimpleDateFormat format_f = new SimpleDateFormat("dd.MM.yyyy");
                        Date dayWeek_f = null;
                        dayWeek_f = format_f.parse(date_fin);
                        // Вывод дня недели даты на экран
                        day_f = new SimpleDateFormat("EEEE").format(dayWeek_f);

                        //System.out.println(((G3Person) g_currPerson).$3N_Ord + "G3Person " + day_f);

                        File file = new File("C:\\JAVA_EXEL\\ORDER_EXEL\\"  + ((G3Person) g_currPerson).$3N_Ord + "_вдк_" +((G3Person) g_currPerson).$3date_Ord_y + ".xls");
                        if (file.exists()) {
                            POIFSFileSystem fs_v = new POIFSFileSystem(new FileInputStream("C:\\JAVA_EXEL\\ORDER_EXEL\\"  + ((G3Person) g_currPerson).$3N_Ord + "_вдк_" +((G3Person) g_currPerson).$3date_Ord_y + ".xls"));
                            Workbook workbook_templ_v = new HSSFWorkbook(fs_v, true);

                            Sheet sheet_templ_v = workbook_templ_v.getSheet("v");

                            if (day_s.contains("воскресенье") || day_s.contains("суббота") || day_f.contains("воскресенье") || day_f.contains("суббота")){
                                //Доступ к скопированному шаблону

                                //Запись общего массива (ФИО, Приказ)
                                _order.add(String.valueOf(((G3Person) g_currPerson).$3N_Ord) + "_вдк");
                                _fio.add(((G3Person) g_currPerson).$3name);

                                //номер приказа
                                int _$N_Ord_v = ((G3Person) g_currPerson).$3N_Ord;
                                Row row_N_Ord_v = sheet_templ_v.getRow(8);
                                Cell cell_N_Ord_v = row_N_Ord_v.getCell(10);
                                cell_N_Ord_v.setCellValue(_$N_Ord_v);

                                //дата приказа
                                String _$date_Ord_v = ((G3Person) g_currPerson).$3date_Ord;
                                Row row_date_Ord_v = sheet_templ_v.getRow(8);
                                Cell cell_date_Ord_v = row_date_Ord_v.getCell(1);
                                cell_date_Ord_v.setCellValue(_$date_Ord_v);

                                //номер записи
                                Row _n_v = sheet_templ_v.getRow(18);
                                Cell _cell_n_v = _n_v.getCell(0);
                                //System.out.println(_cell_n_v.getColumnIndex() + " " + _cell_n_v.getCellType());
                                String n_n_v = _cell_n_v.getStringCellValue();

                                Row _nn_v = sheet_templ_v.getRow(19);
                                Cell _cell_nn_v = _nn_v.getCell(0);
                                //System.out.println(_cell_n_v.getColumnIndex() + " " + _cell_n_v.getCellType());
                                String n_nn_v = _cell_nn_v.getStringCellValue();

                                if(n_n_v.equals("1.") && n_nn_v.equals("2."))  {
                                    String n = "3.";
                                    Row n_v = sheet_templ_v.getRow(20);
                                    Cell cell_n_v = n_v.getCell(0);
                                    cell_n_v.setCellValue(n);

                                    //основная запись
                                    String line = "";
                                    if ((day_s.contains("вос") || day_s.contains("суб")) && (day_f.contains("пон") || day_f.contains("вт")  || day_f.contains("ср")  || day_f.contains("чет")  || day_f.contains("пят"))){
                                        line = ((G3Person) g_currPerson).$3Present_pos + " " + ((G3Person) g_currPerson).$3S_Unit + " " + ((G3Person) g_currPerson).$3name +
                                                " " + "выехать в командировку в " + ((G3Person) g_currPerson).$3adres_city + " в выходной день (" + day_s + ") " +
                                                ((G3Person) g_currPerson).$3date_s_d + " " + mon[Integer.parseInt(((G3Person) g_currPerson).$3date_s_m) - 1] + " 20" + ((G3Person) g_currPerson).$3date_s_y + "г.";
                                        Row line_v = sheet_templ_v.getRow(20);
                                        Cell cell_line_v = line_v.getCell(1);
                                        cell_line_v.setCellValue(line);}

                                    else if ((day_s.contains("вос") || day_s.contains("суб")) && (day_f.contains("вос") || day_f.contains("суб"))){
                                        line = ((G3Person) g_currPerson).$3Present_pos + " " + ((G3Person) g_currPerson).$3S_Unit + " " + ((G3Person) g_currPerson).$3name +
                                                " " + "выехать в командировку в " + ((G3Person) g_currPerson).$3adres_city + " в выходной день (" + day_s + ") " +
                                                ((G3Person) g_currPerson).$3date_s_d + " " + mon[Integer.parseInt(((G3Person) g_currPerson).$3date_s_m) -1] + " 20" + ((G3Person) g_currPerson).$3date_s_y + "г." +
                                                " и вернуться из командировки в " + "выходной день ("+ day_f + ") " + ((G3Person) g_currPerson).$3date_f_d + " " +
                                                mon[Integer.parseInt(((G3Person) g_currPerson).$3date_f_m) -1] + " 20" + ((G3Person) g_currPerson).$3date_f_y + "г.";
                                        Row line_v = sheet_templ_v.getRow(20);
                                        Cell cell_line_v = line_v.getCell(1);
                                        cell_line_v.setCellValue(line);
                                    }
                                    else if ((day_s.contains("пон") || day_s.contains("вт")  || day_s.contains("ср")  || day_s.contains("чет")  || day_s.contains("пят")) && (day_f.contains("вос") || day_f.contains("суб"))){
                                        line = ((G3Person) g_currPerson).$3Present_pos + " " + ((G3Person) g_currPerson).$3S_Unit + " " + ((G3Person) g_currPerson).$3name +
                                                " вернуться из командировки в " + ((G3Person) g_currPerson).$3adres_city + ", в выходной день ("+ day_f + ") " + ((G3Person) g_currPerson).$3date_f_d + " " +
                                                mon[Integer.parseInt(((G3Person) g_currPerson).$3date_f_m) -1] + " 20" + ((G3Person) g_currPerson).$3date_f_y + "г.";
                                        Row line_v = sheet_templ_v.getRow(20);
                                        Cell cell_line_v = line_v.getCell(1);
                                        cell_line_v.setCellValue(line);
                                    }

                                    //с приказом ознакомлен
                                    String end = "С приказом ознакомлен, согласен: ____________  " + ((G3Person) g_currPerson).$3name;
                                    Row end_v = sheet_templ_v.getRow(28);
                                    Cell cell_end_v = end_v.getCell(0);
                                    cell_end_v.setCellValue(end);

                                    //для бухгалтерии
                                    String nn = "4.";
                                    Row nn_v = sheet_templ_v.getRow(21);
                                    Cell cell_nn_v = nn_v.getCell(0);
                                    cell_nn_v.setCellValue(nn);
                                }
                                else if(n_n_v.equals("1."))  {
                                    String n = "2.";
                                    Row n_v = sheet_templ_v.getRow(19);
                                    Cell cell_n_v = n_v.getCell(0);
                                    cell_n_v.setCellValue(n);

                                    //основная запись
                                    String line = "";
                                    if ((day_s.contains("вос") || day_s.contains("суб")) && (day_f.contains("пон") || day_f.contains("вт")  || day_f.contains("ср")  || day_f.contains("чет")  || day_f.contains("пят"))){
                                        line = ((G3Person) g_currPerson).$3Present_pos + " " + ((G3Person) g_currPerson).$3S_Unit + " " + ((G3Person) g_currPerson).$3name +
                                                " " + "выехать в командировку в " + ((G3Person) g_currPerson).$3adres_city + " в выходной день (" + day_s + ") " +
                                                ((G3Person) g_currPerson).$3date_s_d + " " + mon[Integer.parseInt(((G3Person) g_currPerson).$3date_s_m) - 1] + " 20" + ((G3Person) g_currPerson).$3date_s_y + "г.";
                                        Row line_v = sheet_templ_v.getRow(19);
                                        Cell cell_line_v = line_v.getCell(1);
                                        cell_line_v.setCellValue(line);}

                                    else if ((day_s.contains("вос") || day_s.contains("суб")) && (day_f.contains("вос") || day_f.contains("суб"))){
                                        line = ((G3Person) g_currPerson).$3Present_pos + " " + ((G3Person) g_currPerson).$3S_Unit + " " + ((G3Person) g_currPerson).$3name +
                                                " " + "выехать в командировку в " + ((G3Person) g_currPerson).$3adres_city + " в выходной день (" + day_s + ") " +
                                                ((G3Person) g_currPerson).$3date_s_d + " " + mon[Integer.parseInt(((G3Person) g_currPerson).$3date_s_m) -1] + " 20" + ((G3Person) g_currPerson).$3date_s_y + "г." +
                                                " и вернуться из командировки в " + "в выходной день ("+ day_f + ") " + ((G3Person) g_currPerson).$3date_f_d + " " +
                                                mon[Integer.parseInt(((G3Person) g_currPerson).$3date_f_m) -1] + " 20" + ((G3Person) g_currPerson).$3date_f_y + "г.";
                                        Row line_v = sheet_templ_v.getRow(19);
                                        Cell cell_line_v = line_v.getCell(1);
                                        cell_line_v.setCellValue(line);
                                    }
                                    else if ((day_s.contains("пон") || day_s.contains("вт")  || day_s.contains("ср")  || day_s.contains("чет")  || day_s.contains("пят")) && (day_f.contains("вос") || day_f.contains("суб"))){
                                        line = ((G3Person) g_currPerson).$3Present_pos + " " + ((G3Person) g_currPerson).$3S_Unit + " " + ((G3Person) g_currPerson).$3name +
                                                " вернуться из командировки в " + ((G3Person) g_currPerson).$3adres_city + ", в выходной день ("+ day_f + ") " + ((G3Person) g_currPerson).$3date_f_d + " " +
                                                mon[Integer.parseInt(((G3Person) g_currPerson).$3date_f_m) -1] + " 20" + ((G3Person) g_currPerson).$3date_f_y + "г.";
                                        Row line_v = sheet_templ_v.getRow(19);
                                        Cell cell_line_v = line_v.getCell(1);
                                        cell_line_v.setCellValue(line);
                                    }

                                    //с приказом ознакомлен
                                    String end = "С приказом ознакомлен, согласен: ____________  " + ((G3Person) g_currPerson).$3name;
                                    Row end_v = sheet_templ_v.getRow(27);
                                    Cell cell_end_v = end_v.getCell(0);
                                    cell_end_v.setCellValue(end);

                                    //для бухгалтерии
                                    String nn = "3.";
                                    Row nn_v = sheet_templ_v.getRow(21);
                                    Cell cell_nn_v = nn_v.getCell(0);
                                    cell_nn_v.setCellValue(nn);
                                }

                                FileOutputStream fileOut = new FileOutputStream("C:\\JAVA_EXEL\\ORDER_EXEL\\"  + ((G3Person) g_currPerson).$3N_Ord + "_вдк_" +((G3Person) g_currPerson).$3date_Ord_y + ".xls");
                                workbook_templ_v.write(fileOut);
                                fileOut.close();
                                fs_v.close();

                            }

                        }

                        else {
                            POIFSFileSystem fs_v = new POIFSFileSystem(new FileInputStream("C:\\JAVA_EXEL\\generate_v.xls"));
                            Workbook workbook_templ_v = new HSSFWorkbook(fs_v, true);


                            Sheet sheet_templ_v = workbook_templ_v.getSheet("v");

                            if (day_s.contains("вос") || day_s.contains("суб") || day_f.contains("вос") || day_f.contains("суб")){
                                //Доступ к скопированному шаблону

                                //номер приказа
                                int _$N_Ord_v = ((G3Person) g_currPerson).$3N_Ord;
                                Row row_N_Ord_v = sheet_templ_v.getRow(8);
                                Cell cell_N_Ord_v = row_N_Ord_v.getCell(10);
                                cell_N_Ord_v.setCellValue(_$N_Ord_v);

                                //дата приказа
                                String _$date_Ord_v = ((G3Person) g_currPerson).$3date_Ord;
                                Row row_date_Ord_v = sheet_templ_v.getRow(8);
                                Cell cell_date_Ord_v = row_date_Ord_v.getCell(1);
                                cell_date_Ord_v.setCellValue(_$date_Ord_v);

                                //номер записи
                                String n = "1.";
                                Row n_v = sheet_templ_v.getRow(18);
                                Cell cell_n_v = n_v.getCell(0);
                                cell_n_v.setCellValue(n);

                                //основная запись
                                String line = "";
                                if ((day_s.contains("вос") || day_s.contains("суб")) && (day_f.contains("пон") || day_f.contains("вт")  || day_f.contains("ср")  || day_f.contains("чет")  || day_f.contains("пят"))){
                                    line = ((G3Person) g_currPerson).$3Present_pos + " " + ((G3Person) g_currPerson).$3S_Unit + " " + ((G3Person) g_currPerson).$3name +
                                            " " + "выехать в командировку в " + ((G3Person) g_currPerson).$3adres_city + " в выходной день (" + day_s + ") " +
                                            ((G3Person) g_currPerson).$3date_s_d + " " + mon[Integer.parseInt(((G3Person) g_currPerson).$3date_s_m) - 1] + " 20" + ((G3Person) g_currPerson).$3date_s_y + "г.";
                                    Row line_v = sheet_templ_v.getRow(18);
                                    Cell cell_line_v = line_v.getCell(1);
                                    cell_line_v.setCellValue(line);}

                                else if ((day_s.contains("вос") || day_s.contains("суб")) && (day_f.contains("вос") || day_f.contains("суб"))){
                                    line = ((G3Person) g_currPerson).$3Present_pos + " " + ((G3Person) g_currPerson).$3S_Unit + " " + ((G3Person) g_currPerson).$3name +
                                            " " + "выехать в командировку в " + ((G3Person) g_currPerson).$3adres_city + " в выходной день (" + day_s + ") " +
                                            ((G3Person) g_currPerson).$3date_s_d + " " + mon[Integer.parseInt(((G3Person) g_currPerson).$3date_s_m) -1] + " 20" + ((G3Person) g_currPerson).$3date_s_y + "г." +
                                            " и вернуться из командировки в выходной день ("+ day_f + ") " + ((G3Person) g_currPerson).$3date_f_d + " " +
                                            mon[Integer.parseInt(((G3Person) g_currPerson).$3date_f_m) -1] + " 20" + ((G3Person) g_currPerson).$3date_f_y + "г.";
                                    Row line_v = sheet_templ_v.getRow(18);
                                    Cell cell_line_v = line_v.getCell(1);
                                    cell_line_v.setCellValue(line);
                                }
                                else if ((day_s.contains("пон") || day_s.contains("вт")  || day_s.contains("ср")  || day_s.contains("чет")  || day_s.contains("пят")) && (day_f.contains("вос") || day_f.contains("суб"))){
                                    line = ((G3Person) g_currPerson).$3Present_pos + " " + ((G3Person) g_currPerson).$3S_Unit + " " + ((G3Person) g_currPerson).$3name +
                                            " вернуться из командировки в выходной день ("+ day_f + ") " + ((G3Person) g_currPerson).$3date_f_d + " " +
                                            mon[Integer.parseInt(((G3Person) g_currPerson).$3date_f_m) -1] + " 20" + ((G3Person) g_currPerson).$3date_f_y + "г.";
                                    Row line_v = sheet_templ_v.getRow(18);
                                    Cell cell_line_v = line_v.getCell(1);
                                    cell_line_v.setCellValue(line);
                                }

                                //с приказом ознакомлен
                                String end = "С приказом ознакомлен, согласен: ____________  " + ((G3Person) g_currPerson).$3name;
                                Row end_v = sheet_templ_v.getRow(26);
                                Cell cell_end_v = end_v.getCell(0);
                                cell_end_v.setCellValue(end);

                                //для бухгалтерии
                                String nn = "2.";
                                Row nn_v = sheet_templ_v.getRow(21);
                                Cell cell_nn_v = nn_v.getCell(0);
                                cell_nn_v.setCellValue(nn);

                                FileOutputStream fileOut = new FileOutputStream("C:\\JAVA_EXEL\\ORDER_EXEL\\"  + ((G3Person) g_currPerson).$3N_Ord + "_вдк_" +((G3Person) g_currPerson).$3date_Ord_y + ".xls");
                                workbook_templ_v.write(fileOut);
                                fileOut.close();
                                fs_v.close();


                            }

                        }

                        Path file_ = Paths.get(("C:\\JAVA_EXEL\\ORDER_EXEL\\"));
                        BasicFileAttributes attr = Files.readAttributes(file_, BasicFileAttributes.class);
                        // check if the rename operation is success
                        //System.out.println("creationTime: " + attr.creationTime());
                       //System.out.println("lastAccessTime: " + attr.lastAccessTime());
                        //System.out.println("lastModifiedTime: " + attr.lastModifiedTime());
                        //System.out.println("isDirectory: " + attr.isDirectory());
                        //System.out.println("isOther: " + attr.isOther());
                        //System.out.println("isRegularFile: " + attr.isRegularFile());
                        //System.out.println("isSymbolicLink: " + attr.isSymbolicLink());
                        //System.out.println("size: " + attr.size());

                     }
                }
            }
        GruppAllOrderPerson allOrderPerson = new GruppAllOrderPerson();
        for (Map.Entry<String, String> entry : allOrderPerson.gruppAllOrderPerson(_order, _fio).entrySet()) {
            System.out.println("ФИО =  " + entry.getKey() + " № приказа = " + entry.getValue());
            System.out.println();
        }

        }
    }









