package premiumexel;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.IOException;
import java.util.HashMap;

public class GetBasisPrem {

    public HashMap<Integer, String> getBase() throws IOException {

        HashMap<Integer, String> persentbasis = new HashMap<>();
        GetDoc doc = new GetDoc();
        String sheetName = "percent";
        HSSFSheet sheetpers = doc.getSheet(sheetName);
        //System.out.println(sheetName);

        for (int indx = 14; indx < 197; indx++) {
            Row rowpers = sheetpers.getRow(indx - 1);
            Cell cellpers = rowpers.getCell(1);
            int pers = (int) cellpers.getNumericCellValue();
            System.out.println(pers);
            Cell cellbas = rowpers.getCell(2);
            String bas = cellbas.getStringCellValue();
            //System.out.println(bas);
            //bas = bas.replaceAll("+", "");
            persentbasis.put(pers, bas); }
        return persentbasis;
    }
}
