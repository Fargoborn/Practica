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
import java.util.ArrayList;
import java.util.Iterator;

/**
 * @version $Id$
 * @since 0.1
 */

public class ComparatorND {

    long id = 0;
    String name = "";
    String shortname = "";
    int c_yearlast;
    int foodguality;

    C_Docnormative c_docnormativetable = new C_Docnormative();

    public ComparatorND() throws SQLException, ClassNotFoundException {
    }

    public void compare() throws SQLException, IOException {

        for (C_Docnormative c_docnormative : c_docnormativetable.read()) {
            id = c_docnormative.id;
            name = c_docnormative.getName();
            System.out.println(id + " " + name + " " + c_docnormative.getShortname());
        }

        FileInputStream inputStream = new FileInputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\НД\\!НД.xls");
        Workbook doc = new HSSFWorkbook(inputStream);
        Sheet sheet = doc.getSheet("НД");

        Iterator<Row> it = sheet.rowIterator();// Перебираем все строки
        if (it.hasNext()) {
            it.next();
            it.next();
            it.next();
            while (it.hasNext()) {
                Row row = it.next();
                //System.out.println(row.getRowNum());

                //Обрабатываем НД
                for (C_Docnormative c_docnormative : c_docnormativetable.read()) {
                    id = c_docnormative.id;
                    name = c_docnormative.getName().trim();
                    String numberOnly= name.replaceAll("[^0-9]", "");
                    if (c_docnormative.getShortname() != null) {
                        shortname = c_docnormative.getShortname().trim();
                    }
                    Cell nameCell = row.getCell(2);
                    String nameEx = nameCell.getStringCellValue().trim();
                    String numberOnlyEx= nameEx.replaceAll("[^0-9]", "");
                    if (numberOnlyEx.toLowerCase().equals(numberOnly.toLowerCase())) {
                        row.getCell(0).setCellValue(id);
                        row.getCell(1).setCellValue(name);
                    }
                }
            }

            inputStream.close();
            FileOutputStream fileOut = new FileOutputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\НД\\!НД.xls");

            doc.write(fileOut);
            fileOut.close();
            doc.close();
            inputStream.close();

            c_docnormativetable.close();
        }
    }

    public void compareND() throws SQLException, IOException {

        //String comparelineND = "метод";
        String comparelineND = "";
        int idSgm = 220506001;
        String operation = "Insert";

        FileInputStream inputStream = new FileInputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\НД\\!НД.xls");
        Workbook doc = new HSSFWorkbook(inputStream);
        Sheet sheet = doc.getSheet("НД");

        Iterator<Row> it = sheet.rowIterator();// Перебираем все строки
        if (it.hasNext()) {
            while (it.hasNext()) {
                Row row = it.next();
                System.out.println(row.getRowNum());

                //Обрабатываем НД
                Cell cell_nameND = row.getCell(3);
                String nameND = cell_nameND.getStringCellValue();
                Cell cell_shortnameND = row.getCell(4);
                String shortnameND = cell_shortnameND.getStringCellValue().toLowerCase();

                if (nameND.contains("") && shortnameND.contains(comparelineND) && shortnameND.contains("производ")) {
                    row.getCell(0).setCellValue(operation);
                    row.getCell(1).setCellValue(idSgm);
                    idSgm++;
                }
            }

            inputStream.close();
            FileOutputStream fileOut = new FileOutputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\НД\\!НД.xls");

            doc.write(fileOut);
            fileOut.close();
            doc.close();
            inputStream.close();

            c_docnormativetable.close();
        }
    }

    public void idND() throws SQLException, IOException {

        String nameND = "";
        String idLis = "";
        ArrayList<String> list = new ArrayList<>();
        ArrayList<String> list_lis = new ArrayList<>();
        ArrayList<String> list_nameND = new ArrayList<>();


        FileInputStream inputStream = new FileInputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\НД\\!НД.xls");
        Workbook doc = new HSSFWorkbook(inputStream);
        Sheet sheet = doc.getSheet("Заполнен ID");

        Iterator<Row> it = sheet.rowIterator();// Перебираем все строки
        if (it.hasNext()) {
            while (it.hasNext()) {
                Row row = it.next();
                //System.out.println(row.getRowNum());

                //Обрабатываем НД
                Cell cell_nameND = row.getCell(3);
                nameND = cell_nameND.getStringCellValue().trim();
                Cell cell_idND = row.getCell(5);
                int id = (int) cell_idND.getNumericCellValue();
                idLis = "" + id;
                if (id == 0) {
                    idLis = "*";
                }
                list.add(nameND + "&" + idLis);
            }
            inputStream.close();

        FileInputStream inputStream1 = new FileInputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\НД\\!НД.xls");
        Workbook doc1 = new HSSFWorkbook(inputStream1);
            Sheet sheet1 = doc1.getSheet("Классификатор с ID");
            Sheet sheet2 = doc1.getSheet("TypND_NIS");
            Sheet sheet3 = doc1.getSheet("ND_NIS");


            Iterator<Row> it2 = sheet2.rowIterator();// Перебираем все строки
            if (it2.hasNext()) {
                while (it2.hasNext()) {
                    Row row = it2.next();
                    Cell cell_id = row.getCell(0);
                    Cell cell_name = row.getCell(1);
                    String nND = (int) cell_id.getNumericCellValue() + "&" + cell_name.getStringCellValue();
                    //System.out.println(nND);
                    list_nameND.add(nND);
                }
            }

            Iterator<Row> it3 = sheet3.rowIterator();// Перебираем все строки
            String N = "";
            if (it3.hasNext()) {
                while (it3.hasNext()) {
                    Row row = it3.next();
                    Cell cell_id = row.getCell(1);
                    Cell cell_id_nND = row.getCell(0);
                    Cell cell_nameND = row.getCell(2);
                    String id_nND = (int) cell_id_nND.getNumericCellValue() + "";
                    String id = (int) cell_id.getNumericCellValue() + "";

                    for (int i = 0; i < list_nameND.size(); i++) {
                        String line1 = list_nameND.get(i);
                        String[] strings1 = line1.split("&");
                        String nameND1 = strings1[1];
                        String idLis1 = strings1[0];

                        if (cell_nameND.getCellType() == Cell.CELL_TYPE_STRING) {
                            N = cell_nameND.getStringCellValue().trim();
                        }
                        else if (cell_nameND.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            N = (int) cell_nameND.getNumericCellValue() + "";
                        }
                        else N = "NULL";
                        //System.out.println(id_nND);
                        if (id_nND.equals(idLis1) && !N.equals("NULL")) {
                            System.out.println(nameND1 + " " + N + "&" + id);
                            list_lis.add(nameND1.trim() + " " + N + "&" + id);
                        }
                    }
                }
            }


            Iterator<Row> it1 = sheet1.rowIterator();// Перебираем все строки
            if (it1.hasNext()) {
                while (it1.hasNext()) {
                    Row row = it1.next();
                    //System.out.println(row.getRowNum());

                    //Обрабатываем НД
                    Cell cell_nameND = row.getCell(1);
                    String nameNDSGM = cell_nameND.getStringCellValue().trim();

                    for (int i = 0; i < list_lis.size(); i++) {
                        String line = list_lis.get(i);
                        String[] strings = line.split("&");
                        nameND = strings[0].trim();
                        idLis = strings[1].trim();
                        if (nameNDSGM.equals(nameND)) {
                            Cell cell_idND = row.getCell(4);
                            //System.out.println(nameNDSGM);
                            //System.out.println(idLis);
                            cell_idND.setCellValue(idLis);
                        }
                    }

                    for (int i = 0; i < list.size(); i++) {
                        String line = list.get(i);
                        String[] strings = line.split("&");
                        nameND = strings[0].trim();
                        idLis = strings[1].trim();
                        if (nameNDSGM.equals(nameND)) {
                            Cell cell_idND = row.getCell(4);
                            //System.out.println(nameNDSGM);
                            //System.out.println(idLis);
                            cell_idND.setCellValue(idLis);
                        }
                    }
                }
                inputStream1.close();


            FileOutputStream fileOut1 = new FileOutputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\НД\\!НД.xls");

                doc1.write(fileOut1);
                fileOut1.close();
                doc1.close();
                inputStream1.close();

            c_docnormativetable.close();
        }
    }
    }

    public void setINDX() throws SQLException, IOException {

        ArrayList<String> list_name = new ArrayList<>();
        ArrayList<Integer> list_id = new ArrayList<>();

        String nameND = "";

        FileInputStream inputStream = new FileInputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\НД\\!НД.xls");
        Workbook doc = new HSSFWorkbook(inputStream);
        Sheet sheet = doc.getSheet("ND_объем_СГМ");

        Iterator<Row> it = sheet.rowIterator();// Перебираем все строки
        if (it.hasNext()) {
            while (it.hasNext()) {
                Row row = it.next();
                Cell cell_nameND = row.getCell(1);
                if (cell_nameND != null) {
                    nameND = cell_nameND.getStringCellValue().trim();
                }
                Cell cell_Lis_ID = row.getCell(3);
                int id_lis =  (int) cell_Lis_ID.getNumericCellValue();

                list_name.add(nameND);
                list_id.add(id_lis);
            }
        }
            inputStream.close();

            FileInputStream inputStream1 = new FileInputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\НД\\!НД.xls");
            Workbook doc1 = new HSSFWorkbook(inputStream1);
            Sheet sheet1 = doc1.getSheet("Классификатор с ID");

            Iterator<Row> it1 = sheet1.rowIterator();// Перебираем все строки
            if (it1.hasNext()) {
                it1.next();
                it1.next();
                while (it1.hasNext()) {
                    Row row = it1.next();
                    //System.out.println(row.getRowNum());

                    Cell cell_nameND = row.getCell(1);
                    nameND = cell_nameND.getStringCellValue().trim();
                    Cell cell_shortname = row.getCell(2);
                    String shortname = cell_shortname.getStringCellValue().toLowerCase();
                    Cell cell_idND = row.getCell(4);
                    String id = cell_idND.getStringCellValue();
                    System.out.println(id);

                    if (shortname.contains("метод") && !shortname.contains("проб")) {
                        Cell cell_INDX = row.getCell(7);
                        cell_INDX.setCellValue(3);
                    }

                    if (shortname.contains("отбор") && shortname.contains("проб")) {
                        Cell cell_INDX = row.getCell(6);
                        cell_INDX.setCellValue(2);
                    }

                    for (String name : list_name) {
                        if (nameND.equals(name)) {
                            Cell cell_INDX = row.getCell(5);
                            cell_INDX.setCellValue(1);
                        }
                    }

                    int id_lis = 0;
                    for (Integer id_indx : list_id) {
                        if (!id.equals("")) {
                        id_lis = Integer.parseInt(id);
                        }
                        if (id_lis > 0 && id_lis == id_indx) {
                            Cell cell_INDX = row.getCell(5);
                            cell_INDX.setCellValue(1);
                        }
                    }



                }
            }
                inputStream.close();

        FileOutputStream fileOut1 = new FileOutputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\НД\\!НД.xls");

        doc1.write(fileOut1);
        fileOut1.close();
        doc1.close();
        inputStream1.close();

        c_docnormativetable.close();
    }

    public void setINDX_Inggr() throws SQLException, IOException {


        ArrayList<String> list_id = new ArrayList<>();
        ArrayList<String> nameIng = new ArrayList<>();
        String nameING = "";

        FileInputStream inputStream = new FileInputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\Показатели\\Результаты\\Показатель_доработка.xls");
        Workbook doc = new HSSFWorkbook(inputStream);
        Sheet sheet = doc.getSheet("2");

        Iterator<Row> it = sheet.rowIterator();// Перебираем все строки
        if (it.hasNext()) {
            while (it.hasNext()) {
                Row row = it.next();
                Cell cell_nameND = row.getCell(1);
                if (cell_nameND != null) {
                    nameING = cell_nameND.getStringCellValue().toLowerCase().trim().replaceAll("[^a-zA-Zа-яА-Я0-9]", "");
                }
                Cell cell_Lis_ID = row.getCell(0);
                int id_lis = (int) cell_Lis_ID.getNumericCellValue();
                list_id.add(id_lis + "&" + nameING);
            }
        }
        inputStream.close();

        FileInputStream inputStream1 = new FileInputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\Показатели\\Результаты\\Показатель_доработка.xls");
        Workbook doc1 = new HSSFWorkbook(inputStream1);
        Sheet sheet1 = doc1.getSheet("1");
        Sheet sheet2 = doc1.getSheet("2");

        Iterator<Row> it1 = sheet1.rowIterator();// Перебираем все строки
        if (it1.hasNext()) {
            it1.next();
            it1.next();
            while (it1.hasNext()) {
                Row row = it1.next();
                System.out.println(row.getRowNum());

                Cell cell_nameND = row.getCell(1);
                nameING = cell_nameND.getStringCellValue().toLowerCase().trim().replaceAll("[^a-zA-Zа-яА-Я0-9]", "");
                System.out.println(nameING);
                Cell cell_shortname = row.getCell(2);
                String shortname = cell_shortname.getStringCellValue().toLowerCase().trim().replaceAll("[^a-zA-Zа-яА-Я0-9]", "");
                System.out.println(shortname);
                nameIng.add(nameING + "&&" + shortname);

                String id_l = "";
                String name_ing = "";
                for (String id : list_id) {
                    String [] id_ing = id.split("&");
                    id_l = id_ing[0];
                    name_ing = id_ing[1];
                    if (name_ing.equals(nameING) || name_ing.equals(shortname)) {
                        Cell cell_INDX = row.getCell(6);
                        cell_INDX.setCellValue(id_l);
                    }
                }
            }
        }

        Iterator<Row> it2 = sheet2.rowIterator();// Перебираем все строки
        if (it2.hasNext()) {
            while (it2.hasNext()) {
                Row row = it2.next();
                System.out.println(row.getRowNum());

                Cell cell_nameND = row.getCell(1);
                nameING = cell_nameND.getStringCellValue().toLowerCase().trim().replaceAll("[^a-zA-Zа-яА-Я0-9]", "");
                System.out.println(nameING);

                String n = "";
                String sn = "";
                for (String name : nameIng) {
                    String [] n_ing = name.split("&&");
                    if (n_ing.length > 1) {
                    n = n_ing[0].toLowerCase().trim().replaceAll("[^a-zA-Zа-яА-Я0-9]", "");
                    sn = n_ing[1].toLowerCase().trim().replaceAll("[^a-zA-Zа-яА-Я0-9]", "");
                    } else {
                        n = n_ing[0];
                        sn = "$$$$";
                    }
                    if (n.equals(nameING) || sn.equals(nameING)) {
                        row.createCell(2);
                        Cell cell_INDX = row.getCell(2);
                        cell_INDX.setCellValue("$$$");
                    }
                }
            }
        }

        inputStream1.close();

        FileOutputStream fileOut1 = new FileOutputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\Показатели\\Результаты\\Показатель_доработка.xls");

        doc1.write(fileOut1);
        fileOut1.close();
        doc1.close();
        inputStream1.close();

        c_docnormativetable.close();
    }



    public void compareND_dubl() throws SQLException, IOException {

        String comparelineND = "в части";
        //int idSgm = 220506001;
        String operation = "del";
        ArrayList<String> list_nd = new ArrayList<>();
        ArrayList<String> list_nd_lis = new ArrayList<>();
        ArrayList<String> list_nd_res = new ArrayList<>();

        FileInputStream inputStream = new FileInputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\НД_дубли\\без дублей_Классификатор_НД.xls");
        Workbook doc = new HSSFWorkbook(inputStream);
        Sheet sheet = doc.getSheet("1");
        Sheet sheet_del = doc.getSheet("c ЛИС");
        Sheet sheet_r = doc.getSheet("рез");
        Sheet S_del = doc.getSheet("del");


        Iterator<Row> it_res = S_del.rowIterator();// Перебираем все строки
        if (it_res.hasNext()) {
            it_res.next();
            while (it_res.hasNext()) {
                Row row = it_res.next();
                System.out.println(row.getRowNum());

                //Обрабатываем НД
                Cell cell_id = row.getCell(1);
                String id = cell_id.getStringCellValue();
                list_nd_res.add(id);
            }
        }

        /*
        Iterator<Row> it = S_del.rowIterator();// Перебираем все строки
        if (it.hasNext()) {
            it.next();
            while (it.hasNext()) {
                Row row = it.next();

                //Обрабатываем НД
                Cell cell_id_nd = row.getCell(0);
                int id_nd = (int) cell_id_nd.getNumericCellValue();

                for (String id : list_nd_res) {
                String idnd = "" + id_nd;
                if (idnd.equals(id)) {
                row.getCell(5).setCellValue(operation);
                }
                }
            }
        }
        */

        Iterator<Row> it = sheet.rowIterator();// Перебираем все строки
        if (it.hasNext()) {
            it.next();
            while (it.hasNext()) {
                Row row = it.next();
                System.out.println(row.getRowNum());

                //Обрабатываем НД
                Cell cell_id_nd = row.getCell(0);
                int id_nd = (int) cell_id_nd.getNumericCellValue();
                Cell cell_nameND = row.getCell(1);
                String nameND = cell_nameND.getStringCellValue().toLowerCase().trim();
                nameND = nameND.replaceAll("[^0-9]", "");

                //if (nameND.contains(comparelineND)) {
                    //row.getCell(5).setCellValue(operation);
                //}
                //list_nd.add(id_nd + "&&" + nameND);

                for (String id : list_nd_res) {
                    id = id.replaceAll("[^0-9]", "");

                    if (nameND.equals(id)) {
                        System.out.println(id + "###" + nameND);
                    row.getCell(5).setCellValue(operation);
                    }
                }
            }
        }


        /*
        Iterator<Row> it_del = sheet_del.rowIterator();// Перебираем все строки
            if (it_del.hasNext()) {
                while (it_del.hasNext()) {
                    Row row = it_del.next();
                    System.out.println(row.getRowNum());

                    //Обрабатываем НД
                    Cell cell_nameND = row.getCell(1);
                    String nameND = cell_nameND.getStringCellValue().toLowerCase().trim();
                    list_nd_lis.add(nameND);
                }
            }

            int indx = 0;
            for (String str : list_nd) {
                String[] strings = str.split("&&");
                String id = strings[0];
                String name = strings[1];
                sheet_r.createRow(indx);
                Row row = sheet_r.getRow(indx);
                row.createCell(0);
                row.createCell(1);
                row.createCell(2);
                Cell cell_id = row.getCell(0);
                Cell cell_name = row.getCell(1);
                Cell cell_r = row.getCell(2);
                cell_id.setCellValue(id);
                cell_name.setCellValue(name);
                    for (String str_l : list_nd_lis) {
                        String name_l = str_l.replaceAll("[^0-9]", "");
                        //System.out.println(name_l + "   ****  " + str_l);
                        String name_s = name.replaceAll("[^0-9]", "");
                        if (name_s.equals(name_l) && name_l.length() > 1) {
                            cell_r.setCellValue(str_l);
                            System.out.println(name + "   ****  " + str_l);
                            indx++;
                            sheet_r.createRow(indx);
                            row = sheet_r.getRow(indx);
                            row.createCell(2);
                            cell_r = row.getCell(2);
                        }
                    }
                indx++;
            }

            */


            inputStream.close();
            FileOutputStream fileOut = new FileOutputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\НД_дубли\\без дублей_Классификатор_НД.xls");

            doc.write(fileOut);
            fileOut.close();
            doc.close();
            inputStream.close();

            c_docnormativetable.close();
        }

    public void compareUM() throws SQLException, IOException {

        String sgm_um = "";
        String lis_um = "";
        int id_lis = 0;
        ArrayList<String> list_lis = new ArrayList<>();


        FileInputStream inputStream = new FileInputStream("C:\\Users\\dbarshhevskijj\\Desktop\\125792\\Единица измерения\\ЕИ_СГМ.xls");
        Workbook doc = new HSSFWorkbook(inputStream);
        Sheet sheet = doc.getSheet("СГМ");
        Sheet sheet_l = doc.getSheet("ЛИС");

        Iterator<Row> it_res = sheet_l.rowIterator();// Перебираем все строки
        System.out.println("LIS");
        if (it_res.hasNext()) {
            it_res.next();
            while (it_res.hasNext()) {
                Row row = it_res.next();
                System.out.println(row.getRowNum());

                //Обрабатываем НД
                Cell cell_id = row.getCell(2);
                Cell cell_name = row.getCell(1);
                int id = (int) cell_id.getNumericCellValue();
                String name = cell_name.getStringCellValue();
                list_lis.add(id + "&&" + name.toLowerCase().trim());
            }
        }

        Iterator<Row> it = sheet.rowIterator();// Перебираем все строки
        System.out.println("SGM");
        if (it.hasNext()) {
            it.next();
            while (it.hasNext()) {
                Row row = it.next();
                System.out.println(row.getRowNum());

                //Обрабатываем НД
                Cell cell_name = row.getCell(2);
                String name = cell_name.getStringCellValue().toLowerCase().trim();

                String id_l = "";
                String name_l = "";
                for (String str : list_lis) {
                    String [] strings = str.split("&&");
                    id_l = strings[0];
                    name_l = strings[1].toLowerCase().trim();
                    System.out.println(name + "   &&&   " + name_l);
                    if (name.equals(name_l)) {
                        Cell cell_INDX = row.getCell(4);
                        cell_INDX.setCellValue(id_l);
                        Cell cell_n = row.getCell(5);
                        cell_n.setCellValue(name_l);
                    }
                }
            }

            inputStream.close();
            FileOutputStream fileOut = new FileOutputStream("C:\\Users\\dbarshhevskijj\\Desktop\\125792\\Единица измерения\\ЕИ_СГМ.xls");

            doc.write(fileOut);
            fileOut.close();
            doc.close();
            inputStream.close();

            c_docnormativetable.close();
        }
    }

    public void compareDEL() throws SQLException, IOException {

        ArrayList<String> list_del = new ArrayList<>();


        FileInputStream inputStream = new FileInputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\НД_дубли\\без дублей_Классификатор_НД.xls");
        Workbook doc = new HSSFWorkbook(inputStream);
        Sheet sheet = doc.getSheet("Классификатор с ID без дублей");
        Sheet sheet_l = doc.getSheet("Удалить из ЕИАС");
        Sheet sheet_del = doc.getSheet("del");

        /*
        Iterator<Row> it_res = sheet_del.rowIterator();// Перебираем все строки
        System.out.println("del");
        if (it_res.hasNext()) {
            while (it_res.hasNext()) {
                Row row = it_res.next();
                System.out.println(row.getRowNum());

                //Обрабатываем НД
                Cell cell_name = row.getCell(1);
                Cell cell_name_l = row.getCell(2);
                String name = cell_name.getStringCellValue();
                String name_l = cell_name_l.getStringCellValue();
                list_del.add(name.toLowerCase().trim() + "&&" + name_l.toLowerCase().trim());
            }
        }

        Iterator<Row> it = sheet_l.rowIterator();// Перебираем все строки
        System.out.println("sheet_l");
        if (it.hasNext()) {
            while (it.hasNext()) {
                Row row = it.next();
                System.out.println(row.getRowNum());

                //Обрабатываем НД
                Cell cell_name = row.getCell(1);
                String name = cell_name.getStringCellValue().toLowerCase().trim();

                for (String name_del : list_del) {
                    String [] strings = name_del.split("&&");
                    String name_ = strings[0];
                    String name_l = strings[1].toLowerCase().trim();

                    if (name.equals(name_)) {
                        System.out.println(name_l);
                        Cell cell_n = row.getCell(3);
                        cell_n.setCellValue(name_l);
                    }
                }
            }
            */
            Iterator<Row> it_res = sheet.rowIterator();// Перебираем все строки
            if (it_res.hasNext()) {
                it_res.next();
                while (it_res.hasNext()) {
                    Row row = it_res.next();
                    System.out.println(row.getRowNum());
                    //Обрабатываем НД
                    Cell cell_name = row.getCell(1);
                    String name = cell_name.getStringCellValue();
                    Cell cell_id = row.getCell(0);
                    String id = "" + (int) cell_id.getNumericCellValue();
                    list_del.add(id + "&&" + name.toLowerCase().trim());
                }
            }

        Iterator<Row> it = sheet_l.rowIterator();// Перебираем все строки
        System.out.println("sheet_l");
        if (it.hasNext()) {
            while (it.hasNext()) {
                Row row = it.next();
                System.out.println(row.getRowNum());
                System.out.println("***********************");
                //Обрабатываем НД
                Cell cell_name = row.getCell(3);
                String name = cell_name.getStringCellValue().toLowerCase().trim();

                for (String name_del : list_del) {
                    String [] strings = name_del.split("&&");
                    String id = strings[0];
                    String name_ = strings[1].toLowerCase().trim();

                    if (name.equals(name_)) {
                        System.out.println(name_);
                        Cell cell_i = row.getCell(4);
                        cell_i.setCellValue(id);
                        Cell cell_n = row.getCell(5);
                        cell_n.setCellValue(name_);

                    }
                    }

                        /*
                        String nam = name.replaceAll("[^0-9]", "");
                        if (name.contains("замен")) {
                            nam = name.replaceAll(".*?\\(|\\).*?", "").replaceAll("[^0-9]", "");
                        }

                        for (int i = 0; i < list_del.size(); i++) {
                            String [] stringss = list_del.get(i).split("&&");
                            String id_ = stringss[0];
                            String name_1 = stringss[1].toLowerCase().trim();
                            String n = name_1.replaceAll("[^0-9]", "");
                            //System.out.println(i);
                            //System.out.println(nam + " * " + n);
                                if (nam.equals(n)) {
                                    System.out.println(name_1);
                                Cell cell_n = row.getCell(5);
                                cell_n.setCellValue(name_1);
                                Cell cell_i = row.getCell(4);
                                cell_i.setCellValue(id_);
                                break;
                                }
                        }
                        */
                }
            }

            inputStream.close();
            FileOutputStream fileOut = new FileOutputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\НД_дубли\\без дублей_Классификатор_НД.xls");

            doc.write(fileOut);
            fileOut.close();
            doc.close();
            inputStream.close();

            c_docnormativetable.close();
        }


        //Продукты

    private static boolean isDigit(String string) throws NumberFormatException {
        try {
            Integer.parseInt(string);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    public void compareProd() throws SQLException, IOException {

        ArrayList<String> list_prod = new ArrayList<>();
        ArrayList<String> list_prod_p = new ArrayList<>();

        FileInputStream inputStream = new FileInputStream("C:\\Users\\dbarshhevskijj\\Desktop\\125792\\Продукт\\Prod.xls");
        Workbook doc = new HSSFWorkbook(inputStream);
        Sheet sheet = doc.getSheet("СГМ_прод");
        Sheet sheet_pp = doc.getSheet("SGM_Prod_prim");
        Sheet sheet_l = doc.getSheet("ЛИС");
        Sheet sheet_rm = doc.getSheet("Рабочее место_помещение_СГМ");

        //Продукт
        Iterator<Row> it_res = sheet.rowIterator();// Перебираем все строки
        System.out.println("SGM");
        if (it_res.hasNext()) {
            it_res.next();
            while (it_res.hasNext()) {
                Row row = it_res.next();
                System.out.println(row.getRowNum());

                Cell cell_id = row.getCell(0);
                Cell cell_name = row.getCell(1);
                long id = (long) cell_id.getNumericCellValue();
                String name = cell_name.getStringCellValue();
                list_prod.add(id + "&&" + name.toLowerCase().trim());
            }
        }

        //Продукт примечание
        Iterator<Row> it_pp = sheet_pp.rowIterator();// Перебираем все строки
        System.out.println("SGM_Prod_prim");
        if (it_pp.hasNext()) {
            it_pp.next();
            while (it_pp.hasNext()) {
                Row row = it_pp.next();
                System.out.println(row.getRowNum());

                Cell cell_id = row.getCell(0);
                Cell cell_name = row.getCell(1);
                long id = (long) cell_id.getNumericCellValue();
                String name = cell_name.getStringCellValue();
                list_prod_p.add(id + "&&" + name.toLowerCase().trim());
            }
        }

        Iterator<Row> itt = sheet_l.rowIterator();// Перебираем все строки
        System.out.println("LIS");
        if (itt.hasNext()) {
            itt.next();
            while (itt.hasNext()) {
                Row row = itt.next();
                System.out.println(row.getRowNum());

                //Обрабатываем НД
                Cell cell_name = row.getCell(2);
                String name = cell_name.getStringCellValue().toLowerCase().trim();

                while (true) {
                    if (ComparatorND.isDigit(name.substring(0 , 1))) {
                        name = name.substring(1 , name.length());
                    } else {
                        break;
                    }
                }
                cell_name.setCellValue(name);
            }
        }

        Iterator<Row> it = sheet_l.rowIterator();// Перебираем все строки
        System.out.println("LIS");
        if (it.hasNext()) {
            it.next();
            while (it.hasNext()) {
                Row row = it.next();
                System.out.println(row.getRowNum());

                //Обрабатываем НД
                Cell cell_id = row.getCell(1);
                String id = "" + (int) cell_id.getNumericCellValue();
                Cell cell_name = row.getCell(2);
                String name = cell_name.getStringCellValue().toLowerCase().trim().replaceAll("[^А-Яа-я]", "");
                String name_ppl = cell_name.getStringCellValue().toLowerCase().trim();
                String id_l = "";
                String name_l = "";
                Cell cell_n = row.getCell(4);
                Cell cel_id_sgm = row.getCell(3);
                cell_n.setCellValue("!!");
                Cell cel_id_sgm_pp = row.getCell(5);
                Cell cell_pp = row.getCell(6);
                cell_pp.setCellValue("!!");
                for (String str : list_prod) {
                    String [] strings = str.split("&&");
                    id_l = strings[0];
                    name_l = strings[1].toLowerCase().trim();
                    String n = name_l.replaceAll("[^А-Яа-я]", "");
                    //System.out.println(name + "   &&&   " + name_l);
                    if (name.equals(n)) {
                        cel_id_sgm.setCellValue(id_l);
                        cell_n.setCellValue(name_l);
                    } else if (name.length() > n.length()) {
                        if (name.contains(n) && n.length() > cell_n.getStringCellValue().length()) {
                            for (String str_pp : list_prod_p) {
                                String [] strings_pp = str_pp.split("&&");
                                String id_pp = strings_pp[0];
                                String name_pp = strings_pp[1].toLowerCase().trim();
                                String pp = name_pp.replaceAll("[^А-Яа-я]", "");
                                //if (name.contains(pp) && pp.length() > cell_pp.getStringCellValue().length()) {
                                //    cell_pp.setCellValue(name_pp);
                                //    cel_id_sgm_pp.setCellValue(id_pp);
                                //}
                                if (name_ppl.contains(name_pp) && name_pp.length() > cell_pp.getStringCellValue().length()) {
                                    System.out.println(name_ppl + "   &&&   " + name_pp);
                                    cell_pp.setCellValue(name_pp);
                                    cel_id_sgm_pp.setCellValue(id_pp);
                                }}
                            cell_n.setCellValue(name_l);
                            cel_id_sgm.setCellValue(id_l);

                    }
                } else if (name.length() < n.length()) {
                    if (n.contains(name) && n.length() > cell_n.getStringCellValue().length()) {
                            cell_n.setCellValue(name_l);
                            cel_id_sgm.setCellValue(id_l);
                    }
                }
            }
            }
        }

            inputStream.close();
            FileOutputStream fileOut = new FileOutputStream("C:\\Users\\dbarshhevskijj\\Desktop\\125792\\Продукт\\Prod.xls");

            doc.write(fileOut);
            fileOut.close();
            doc.close();
            inputStream.close();

            c_docnormativetable.close();

    }



    public void corr_setINDX() throws SQLException, IOException {

        ArrayList<Integer> list_id = new ArrayList<>();

        String nameND = "";

        FileInputStream inputStream1 = new FileInputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\НД_дубли\\доработка_Классификатор_НД.xls");
        Workbook doc1 = new HSSFWorkbook(inputStream1);
        Sheet sheet1 = doc1.getSheet("Классификатор с ID без дублей");
        Sheet sheet_pp = doc1.getSheet("1");

        Iterator<Row> it_pp = sheet_pp.rowIterator();// Перебираем все строки
        if (it_pp.hasNext()) {
            while (it_pp.hasNext()) {
                Row row = it_pp.next();
                System.out.println(row.getRowNum());

                Cell cell_id = row.getCell(1);
                int id = (int) cell_id.getNumericCellValue();
                list_id.add(id);
            }
        }


        Iterator<Row> it1 = sheet1.rowIterator();// Перебираем все строки
        if (it1.hasNext()) {
            it1.next();
            while (it1.hasNext()) {
                Row row = it1.next();
                System.out.println(row.getRowNum());

                Cell cell_id = row.getCell(0);
                int id = (int) cell_id.getNumericCellValue();

                for (int idd : list_id) {
                    if (id == idd) {
                        Cell cell_idd = row.getCell(8);
                        cell_idd.setCellValue("&&");
                    }
                }

                Cell cell_shortname = row.getCell(2);
                String shortname = cell_shortname.getStringCellValue().toLowerCase();
                Cell cell_idND = row.getCell(0);
                //String id = cell_idND.getStringCellValue();
                //System.out.println(id);


                if (shortname.contains("пдк") || shortname.contains("пду")|| shortname.contains("обув") || shortname.contains("нормы") || shortname.contains("гигиенические требования")
                        || shortname.contains("допустимые уровни") || shortname.contains("норматив") || shortname.contains("показатель") || shortname.contains("эпидемиологические требования")
                        || shortname.contains("требования")) {
                    Cell cell_INDX = row.getCell(5);
                    cell_INDX.setCellValue(1);
                }

                if (shortname.contains("измерение") || shortname.contains("методика измерени")) {
                    Cell cell_INDX = row.getCell(7);
                    cell_INDX.setCellValue(3);
                }

                if (shortname.contains("метод") && !shortname.contains("проб") && !shortname.contains("методические")) {
                    Cell cell_INDX = row.getCell(7);
                    cell_INDX.setCellValue(3);
                }

            }
        }

        FileOutputStream fileOut1 = new FileOutputStream("C:\\Users\\dbarshhevskijj\\Desktop\\119834\\119834\\НД_дубли\\доработка_Классификатор_НД.xls");

        doc1.write(fileOut1);
        fileOut1.close();
        doc1.close();
        inputStream1.close();

        c_docnormativetable.close();
    }

    public void work_p() throws SQLException, IOException {

        ArrayList<String> list_id = new ArrayList<>();
        ArrayList<String> list_eom = new ArrayList<>();

        String nameWP = "";

        FileInputStream inputStream1 = new FileInputStream("C:\\Users\\dbarshhevskijj\\Desktop\\125792\\Продукт\\РМ.xls");
        Workbook doc1 = new HSSFWorkbook(inputStream1);
        Sheet sheet_wpsgm = doc1.getSheet("РМ_СГМ");
        Sheet sheet_wp = doc1.getSheet("РМ_П_ЛИС");
        Sheet sheet_eomtype = doc1.getSheet("eomtype");

        Iterator<Row> it_eom = sheet_eomtype.rowIterator();// Перебираем все строки
        if (it_eom.hasNext()) {
            it_eom.next();
            while (it_eom.hasNext()) {
                Row row = it_eom.next();
                System.out.println(row.getRowNum());
                Cell cell_id = row.getCell(0);
                int id = (int) cell_id.getNumericCellValue();
                Cell cell_eom = row.getCell(2);
                String eom = ("" + (int) cell_eom.getNumericCellValue()).toLowerCase().trim();
                String eomWP = id + "&&" + eom;
                list_eom.add(eomWP);
            }
        }

        Iterator<Row> it_pp = sheet_wpsgm.rowIterator();// Перебираем все строки
        if (it_pp.hasNext()) {
            it_pp.next();
            while (it_pp.hasNext()) {
                Row row = it_pp.next();
                System.out.println(row.getRowNum());

                Cell cell_id = row.getCell(0);
                int id = (int) cell_id.getNumericCellValue();
                Cell cell_name = row.getCell(1);
                String name = cell_name.getStringCellValue().toLowerCase().trim();
                nameWP = id + "&&" + name;
                list_id.add(nameWP);
            }
        }

        Iterator<Row> it1 = sheet_wp.rowIterator();// Перебираем все строки
        if (it1.hasNext()) {
            it1.next();
            while (it1.hasNext()) {
                Row row = it1.next();
                System.out.println(row.getRowNum());

                Cell cell_id = row.getCell(1);
                String id = ("" + (int) cell_id.getNumericCellValue()).toLowerCase().trim();
                Cell cell_name = row.getCell(4);
                String name = cell_name.getStringCellValue();

                for (String string : list_id) {
                    String [] strings = string.split("&&");
                    String id_s = strings[0];
                    String name_s = strings[1].toLowerCase().trim();
                    if (name.equals(name_s)) {
                        Cell cell_idd = row.getCell(5);
                        cell_idd.setCellValue(id_s);
                        Cell cell_n = row.getCell(6);
                        cell_n.setCellValue(name_s);
                    }
                }

                for (String string : list_eom) {
                    String [] strings = string.split("&&");
                    String id_s = strings[0];
                    String name_s = strings[1].toLowerCase().trim();
                    //System.out.println(name_s);
                    if (id.equals(id_s)) {
                        Cell cell_idd = row.getCell(2);
                        cell_idd.setCellValue(name_s);
                    }
                }
            }
        }

        FileOutputStream fileOut1 = new FileOutputStream("C:\\Users\\dbarshhevskijj\\Desktop\\125792\\Продукт\\РМ.xls");

        doc1.write(fileOut1);
        fileOut1.close();
        doc1.close();
        inputStream1.close();


    }


    public static void main (String[] args) throws IOException, SQLException, ClassNotFoundException {

        ComparatorND comparatorND = new ComparatorND();
        //comparatorND.compareND();
        //comparatorND.idND();
        //comparatorND.setINDX();
        //comparatorND.setINDX_Inggr();
        //comparatorND.compareND_dubl();
        //comparatorND.compareUM();
        //comparatorND.compareDEL();
        //comparatorND.compareProd();
        //comparatorND.corr_setINDX();
        comparatorND.work_p();
    }
}
