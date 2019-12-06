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

public class Dublicator {

    int prod;
    String prodname;
    int note;
    String notename;
    String exp;
    ArrayList<PDKProductObj> pdkProducts = new ArrayList<>();

    public Dublicator() throws SQLException {
    }

    /**
     * Метод переноса данных из первичной таблицы
     * в таблицу для формирования скриптов.
     * @return
     */
    public ArrayList<PDKProductObj> dublicate() throws IOException, SQLException, ClassNotFoundException {
        FileInputStream inputStream = new FileInputStream("C:\\JAVA_EXEL\\PDKConverter\\normative.xls");
        Workbook doc = new HSSFWorkbook(inputStream);
        Sheet sheetIN = doc.getSheet("1");
        Sheet sheetOUT = doc.getSheet("PdkProduct");
        int indx = 3;
        S_PDKProduct s_pdkProduct = new S_PDKProduct();
        //System.out.println(s_pdkProduct.getMaxID());
        long id = s_pdkProduct.getMaxID();
        s_pdkProduct.close();

        //Читаем НД
        Row row_nd = sheetIN.getRow(1);
        Cell nd_cell = row_nd.getCell(0);
        Cell ndname_cell = row_nd.getCell(1);
        String nd = nd_cell.getStringCellValue();
        String ndname = ndname_cell.getStringCellValue();

        Iterator<Row> it = sheetIN.rowIterator();// Перебираем все строки
        if (it.hasNext()) {
            it.next();
            it.next();
            it.next();
            while (it.hasNext()) {
            Row row = it.next();
            //читаем
                Cell prod_cell = row.getCell(0);
                if (!prod_cell.toString().equals("")) {
                prod = (int) prod_cell.getNumericCellValue();}
                Cell prodname_cell = row.getCell(1);
                if (!prodname_cell.toString().equals("")) {
                prodname = prodname_cell.getStringCellValue().toLowerCase();}
                note = 0;
                Cell note_cell = row.getCell(2);
                if (!note_cell.toString().equals("")) {
                note = (int) note_cell.getNumericCellValue();}
                notename = "";
                Cell notename_cell = row.getCell(3);
                if (!notename_cell.toString().equals("")) {
                notename = notename_cell.getStringCellValue();}
                Cell ing_cell = row.getCell(4);
                int ing = (int) ing_cell.getNumericCellValue();
                Cell ingname_cell = row.getCell(5);
                String ingname = ingname_cell.getStringCellValue();
                Cell unit_cell = row.getCell(7);
                int unit = (int) unit_cell.getNumericCellValue();
                Cell unitname_cell = row.getCell(8);
                String unitname = unitname_cell.getStringCellValue();
                Cell min_cell = row.getCell(11);
                String min = min_cell.getStringCellValue();
                Cell max_cell = row.getCell(12);
                String max = max_cell.getStringCellValue();
                Cell exp_cell = row.getCell(13);
                exp = exp_cell.getStringCellValue();
                String expn[] = exp.split(".0");
                if (expn.length > 0) {
                exp = expn[0];
                }
            //пишем в контрольную таблицу
                sheetOUT.createRow(indx);
                Row rowOut = sheetOUT.getRow(indx);
                for (int i = 0; i < 28; i++) {
                rowOut.createCell(i);
                }
                rowOut.getCell(0).setCellValue(id);
                rowOut.getCell(1).setCellValue(prod);
                rowOut.getCell(2).setCellValue(prodname);
                rowOut.getCell(3).setCellValue(note);
                rowOut.getCell(4).setCellValue(notename);
                rowOut.getCell(7).setCellValue(ing);
                rowOut.getCell(8).setCellValue(ingname);
                rowOut.getCell(9).setCellValue(unit);
                rowOut.getCell(10).setCellValue(unitname);
                rowOut.getCell(11).setCellValue(max);
                rowOut.getCell(12).setCellValue(min);
                rowOut.getCell(13).setCellValue(nd);
                rowOut.createCell(28);
                if (exp != null && !exp.equals("0")) {
                    rowOut.getCell(28).setCellValue(exp);
                }
                //PDKProductObj(long ID, int c_PRODUCT, int c_NOTE, int c_INGREDIENT, int c_UNIT, String MINVAL, String MAXVAL, int c_DOCPDK, int c_LABORATORY)
                PDKProductObj pdkProductObj = new PDKProductObj(id, prod, note, ing, unit, min, max, nd, 0, exp);
                pdkProducts.add(pdkProductObj);
                rowOut.getCell(5).setCellValue(pdkProductObj.getC_PACKING());
                rowOut.getCell(15).setCellValue(pdkProductObj.getC_METHODDEFINE());
                rowOut.getCell(18).setCellValue(pdkProductObj.getC_LABORATORY());
                rowOut.getCell(20).setCellValue(pdkProductObj.getCOUNTRESEARCH());
                rowOut.getCell(22).setCellValue(pdkProductObj.getC_TYPETEST());
                rowOut.getCell(24).setCellValue(pdkProductObj.getNORMTXT());
                rowOut.getCell(25).setCellValue(pdkProductObj.getVALTXTOK());
                rowOut.getCell(26).setCellValue(pdkProductObj.getVALTXTBAD());
                rowOut.getCell(27).setCellValue(pdkProductObj.getC_STANDARD());

                indx++;
                id++;
            }
        }
        inputStream.close();
        FileOutputStream fileOut = new FileOutputStream("C:\\JAVA_EXEL\\PDKConverter\\normative.xls");

        doc.write(fileOut);
        fileOut.close();
        doc.close();
        inputStream.close();

        return pdkProducts;
    }
}
