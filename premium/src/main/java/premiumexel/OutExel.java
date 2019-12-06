package premiumexel;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;

public class OutExel {

    private static File fileout;

    public void copyExel() throws IOException {
        //Копируем шаблон
        File filev = new File(new File("").getAbsolutePath() + "\\Templates\\prem_out.xls");
        FileOutputStream fileOutputStreamv = new FileOutputStream(new File("").getAbsolutePath() + "\\Templates\\generate_prem_out.xls");
        Files.copy(filev.toPath(), fileOutputStreamv);
        Workbook workbookv = new HSSFWorkbook();
        workbookv.write(fileOutputStreamv);
        fileOutputStreamv.close();

        fileout = new File(new File("").getAbsolutePath() + "\\generate_prem_out.xls");

    }

    public HSSFSheet outExelSheet() throws IOException {

        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(OutExel.fileout));
        Workbook workbooktempl = new HSSFWorkbook(fs, true);
        Sheet sheetout = workbooktempl.getSheet("out");

        return (HSSFSheet) sheetout;
    }
}
