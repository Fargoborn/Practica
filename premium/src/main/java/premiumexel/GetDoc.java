package premiumexel;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * @version $Id$
 * @since 0.1
 */
public class GetDoc {

  /**
   * Метод реализует доступ к таблице Exel.
   */

  private Print print = new Print();
  private String[] fn = print.getFileName();
  private String pathoutfile = print.getOutFileName();

  public GetDoc() throws IOException {
  }

  //получение книги для чтения
  public HSSFWorkbook getExelDoc() throws IOException {
    FileInputStream in = null;
    in = new FileInputStream(fn[0]);
    System.out.println(fn[0] + "*");
    HSSFWorkbook wb = null;
    wb = new HSSFWorkbook(in);
    return wb;
  }

  public HSSFWorkbook getExelDocPers() throws IOException {
    FileInputStream in = null;
    in = new FileInputStream(fn[1]);
    //System.out.println(fn[1] + "**");
    HSSFWorkbook wb = null;
    wb = new HSSFWorkbook(in);
    //System.out.println(fn[1] + "**");
    //System.out.println(wb.getActiveSheetIndex());
    return wb;
  }

  //получение книги для записи
  public HSSFWorkbook outExelDoc() throws IOException {
    FileInputStream in = null;
    in = new FileInputStream(pathoutfile);
    HSSFWorkbook wb = null;
    wb = new HSSFWorkbook(in);
    return wb;
  }

  //получение страницы для чтения
  public HSSFSheet getSheet(final String sheetName) throws IOException {
    GetDoc doc = new GetDoc();
    return doc.getExelDoc().getSheet(sheetName);
  }

  public HSSFSheet getSheetPers(final String sheetName) throws IOException {
    GetDoc doc = new GetDoc();
    return doc.getExelDocPers().getSheet(sheetName);
  }

  //получение страницы для записи
  public HSSFSheet getOutSheet(final String sheetName) throws IOException {
  GetDoc outdoc = new GetDoc();
  return outdoc.outExelDoc().getSheet(sheetName);
  }

}

