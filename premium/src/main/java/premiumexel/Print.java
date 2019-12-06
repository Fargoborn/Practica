package premiumexel;

import org.apache.poi.hssf.usermodel.HSSFSheet;
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
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Random;




public class Print {

    //final String FILENAME = "C:\\JAVA_EXEL\\Premium\\prem.xls";
    //final String FILENAME = (new File(".").getAbsolutePath() + "\\prem.xls");

    /**
     * Метод формирует путь к файлу Exel.
     */
    public String[] getFileName() {
        String pathfile = "";
        String pf = "";
        String[] fn = new String[]{};
        String fp = new File("").getAbsolutePath();
        File[] files = new File(fp).listFiles();
        for (File file : files) {
            if (file.getName().contains(".xls") && file.getName().equals("prem.xls")) {
                pathfile = file.getName();
            }
            if (file.getName().contains(".xls") && !file.getName().equals("prem.xls")) {
                pf = file.getName();
            }
            fn = new String[] {pathfile, pf};
        }
        //System.out.println(fp +"\\" + pathfile);
        return fn;
    }

    public String getOutFileName() {
        String pathoutfile = new File("").getAbsolutePath() + "\\Templates\\generate_prem_out.xls";
        //final String FILENAME = (new File(".").getAbsolutePath() + "\\prem.xls");
        return pathoutfile;
    }

    public static void main(String[] args) throws IOException {

        String b5 = "за оперативное выполнение особо срочных заданий руководства - 5%";
        String b7 = "за наставничество молодым специалистам и обучение новых сотрудников - 7%";
        String b10 = "за профессиональное мастерство - 10%";
        String b20 = "за оперативное выполнение особо важных заданий руководства - 20%";
        String b200 = "за оперативное и качественное выполнение задач внутреннего нормоконтроля документации - 20%";
        String b30 = "за активное участие и большой вклад в реализацию контрактов и проектов предприятия - 30%";
        String b40 = "за бездефектное изготовление продукции, отсутствие претензий к качеству продукции (рекламаций) - 40%";
        String b50 = "за обеспечение качества продукции (работ, услуг), соответствующего международным стандартам (ГОСТ ИСО 9001 - 2011) - 50%";
        String b100 = "за руководство процессом проектирования и разработки программного обеспечения - 100%";
        String b1000 = "за личный вклад в реализацию научных, научно-исследовательских и технологических работ, внедрение инновационных подходов  к разработке программного обеспечения - 100%";
        String b10000 = "за использование специализированных технологий, обеспечивающих снижение трудоемкости разработки программного обеспечения и снижение себестоимости продукции - 100%";

        String[] basis = new String[] {"", b5, b7, b10, b20, b200, b30, b40, b50, b100, b1000, b10000};

       GetDoc getExelDoc = new GetDoc();
       System.out.println(getExelDoc.getExelDoc());

        GetDoc doc = new GetDoc();
        //String sheetName = "out";
        HSSFSheet sheetout = doc.getExelDocPers().getSheetAt(0);

        OutExel outExel = new OutExel();
        outExel.copyExel();
        POIFSFileSystem file = new POIFSFileSystem(new FileInputStream(new File("").getAbsolutePath() + "\\Templates\\generate_prem_out.xls"));
        Workbook workbooktempl = new HSSFWorkbook(file, true);
        Sheet outsheet = workbooktempl.getSheet("out");
        int indx = 8;

        GetBasisPrem basisPrem = new GetBasisPrem();
        HashMap persbase = basisPrem.getBase();
        //while (it.hasNext()) {
            //Map.Entry pair = (Map.Entry)it.next();
            //System.out.println(pair.getKey() + " " + pair.getValue());
        //}

        Iterator<Row> rowIterator = sheetout.rowIterator();
        if (rowIterator.hasNext()) {
            rowIterator.next();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                String name = "";
                int namepers = 0;
                if (row.getCell(1).getStringCellValue().equals("")) {
                    break;
                }
                Cell cellname = row.getCell(1);
                name = cellname.getStringCellValue();
                Cell cellnameperc = row.getCell(2);
                namepers = (int) cellnameperc.getNumericCellValue();
                String perc = String.valueOf(namepers);
                //System.out.println(name + " " + namepers);

                Iterator it = persbase.entrySet().iterator();
                while (it.hasNext()) {
                    Map.Entry pair = (Map.Entry) it.next();
                    String perbas = "";
                    String[] list = new String[]{};
                    //System.out.println((int)pair.getKey() + "*" + pair.getValue() + " " + name + " " + namepers);

                  //System.out.println(indx);
                  outsheet.createRow(indx);
                  Row outrow = outsheet.getRow(indx);

                    if (pair.getKey().toString().equals(perc)) {
                        //System.out.println((int)pair.getKey() + "%  " + name + " " + pair.getValue() + "  --" + perc);
                        perbas = pair.getValue().toString().replace(",", "");

                        //System.out.println(per_bas);
                        for (String retval : perbas.split(" ")) {
                            //System.out.println(retval.trim() + "*");
                            retval = retval.trim().replace("+", "W");
                            //System.out.println(retval);
                            list = retval.split(" ");
                        }
                        int rnd = new Random().nextInt(list.length);
                        System.out.println(name + " *****" + namepers + "% " + list[rnd]);
                        String[] listchar = list[rnd].split("W");

                        String out = "";
                        for (String elem : listchar) {
                            String base = "";
                            System.out.println(Integer.parseInt(String.valueOf(elem)));
                            base = basis[Integer.parseInt(String.valueOf(elem))];
                            System.out.println(base);
                                out = out + "; " + base;
                        }
                        System.out.println(name + " " + namepers + "% " + out);

                        outrow.createCell(0);
                        Cell cellFIO = outrow.getCell(0);
                        cellFIO.setCellValue(name);
                        outrow.createCell(1);
                        Cell cellbase = outrow.getCell(1);
                        cellbase.setCellValue(out.substring(2));
                        outrow.createCell(2);
                        Cell cellperc = outrow.getCell(2);
                        cellperc.setCellValue(namepers);
                        break;
                    } else {
                      outrow.createCell(0);
                      Cell cellFIO = outrow.getCell(0);
                      cellFIO.setCellValue(name);
                      outrow.createCell(2);
                      Cell cellperc = outrow.getCell(2);
                      cellperc.setCellValue(namepers);
                    }

                    if (!pair.getKey().toString().equals(perc))  {
                        outsheet.createRow(indx);
                        outrow.createCell(0);
                        Cell cellFIO = outrow.getCell(0);
                        cellFIO.setCellValue(name);
                        outrow.createCell(1);
                        Cell cellbase = outrow.getCell(1);
                        cellbase.setCellValue("----------->>>> КОМБИНАЦИЯ НЕ НАЙДЕНА <<<<-------------");
                        outrow.createCell(2);
                        Cell cellperc = outrow.getCell(2);
                        cellperc.setCellValue(namepers);
                    }
                }
                indx++;
            }
        }
        System.out.println("***");
        System.out.println("УРА!!!");
        System.out.println("УРА!!!");
        System.out.println("УРА!!!");
        Print p = new Print();
        try (FileOutputStream fileOut = new FileOutputStream(new File("").getAbsolutePath() + "\\generate_" + p.getFileName()[1])) {
            ((HSSFWorkbook) workbooktempl).write(fileOut);
            fileOut.close();
        }
        file.close();
        workbooktempl.close();
    }
}
