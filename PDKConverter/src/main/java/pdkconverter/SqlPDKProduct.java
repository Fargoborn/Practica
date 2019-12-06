package pdkconverter;
import java.io.*;
import java.util.ArrayList;


public class SqlPDKProduct {

    SqlPDKProduct () {
    }

    private Integer ID = null;
    private Integer C_PRODUCT = null;
    private Integer C_NOTE = null;
    private Integer C_PACKING = null;
    private Integer C_INGREDIENT = null;
    private Integer C_UNIT = null;
    private Double MAXVAL = null;
    private Double MINVAL = null;
    private Integer C_DOCPDK = null;
    private Integer C_METHODDEFINE = null;
    private String DOCNORMATIVENAME = "null";
    private Integer C_LABORATORY = null;
    private Integer COUNTRESEARCH = null;
    private String NOTE = "null";
    private Integer C_TYPETEST = null;
    private String NORMTXT = "null";
    private String VALTXTOK = "null";
    private String VALTXTBAD = "null";
    private Integer C_STANDART = null;
    private String Exp = "";

    public void mains()throws Exception {

        /**
         * Метод формирует файл со скриптом на вставку записей по Продуктам.
         * Код писался, когда я еще не умел работать с объектами, поэтому столько массивов :))
         * Немного доработан, но полностью переделывать не стал.
         */

        ArrayList<Integer> list_ID = new ArrayList<>();
        ArrayList<Integer> list_C_PRODUCT = new ArrayList<>();
        ArrayList<Integer> list_C_NOTE = new ArrayList<>();
        ArrayList<Integer> list_C_PACKING = new ArrayList<>();
        ArrayList<Integer> list_C_INGREDIENT = new ArrayList<>();
        ArrayList<Integer> list_C_UNIT = new ArrayList<>();
        ArrayList<Double> list_MAXVAL = new ArrayList<>();
        ArrayList<Double> list_MINVAL = new ArrayList<>();
        ArrayList<Integer> list_C_DOCPDK = new ArrayList<>();
        ArrayList<Integer> list_C_METHODDEFINE = new ArrayList<>();
        ArrayList<String> list_DOCNORMATIVENAME = new ArrayList<>();
        ArrayList<Integer> list_C_LABORATORY = new ArrayList<>();
        ArrayList<Integer> list_COUNTRESEARCH = new ArrayList<>();
        ArrayList<String> list_NOTE = new ArrayList<>();
        ArrayList<Integer> list_C_TYPETEST = new ArrayList<>();
        ArrayList<String> list_NORMTXT = new ArrayList<>();
        ArrayList<String> list_VALTXTOK = new ArrayList<>();
        ArrayList<String> list_VALTXTBAD = new ArrayList<>();
        ArrayList<Integer> list_C_STANDART = new ArrayList<>();
        ArrayList<String> list_Exp = new ArrayList<>();
        ArrayList<String> arrayList_SQL = new ArrayList<>();

            System.out.println("*****");
            Dublicator dublicator = new Dublicator();
            ArrayList<PDKProductObj> pdkProductObjs = dublicator.dublicate();
            for (PDKProductObj pdkProductObj : pdkProductObjs) {
                try{
                    ID = (int) pdkProductObj.getID();
                    System.out.println(ID + " ID");
                    C_PRODUCT = pdkProductObj.getC_PRODUCT();
                    System.out.println(C_PRODUCT);
                    C_NOTE = pdkProductObj.getC_NOTE();
                    System.out.println(C_NOTE);
                    C_PACKING = pdkProductObj.getC_PACKING();
                    System.out.println(C_PACKING);
                    C_INGREDIENT = pdkProductObj.getC_INGREDIENT();
                    System.out.println(C_INGREDIENT);
                    C_UNIT = pdkProductObj.getC_UNIT();
                    System.out.println(C_UNIT);

                    //NORMTXT = NORMTXT_Cell.getStringCellValue();
                    if (pdkProductObj.getNORMTXT() == null){
                        NORMTXT = "null";
                        System.out.println(NORMTXT + " = null");
                    }
                    else{
                        NORMTXT = pdkProductObj.getNORMTXT();
                        System.out.println(NORMTXT);
                    }

                    //System.out.println(" MAXVAL_CellType " + MAXVAL_Cell.getCellType());
                    if (!pdkProductObj.getMAXVAL().equals("")) {
                    MAXVAL = Double.parseDouble(pdkProductObj.getMAXVAL().replace(",", "."));
                    if (MAXVAL == 0.0) {
                        NORMTXT = "не допускается";
                    }
                    System.out.println(MAXVAL + " MAXVAL");
                    } else {
                        MAXVAL = null;
                    }
                    //System.out.println(" MINVAL_CellType" + MINVAL_Cell.getCellType());
                    if (!pdkProductObj.getMINVAL().equals("")) {
                    MINVAL = Double.parseDouble(pdkProductObj.getMINVAL().replace(",", "."));
                    System.out.println(MINVAL + " MINVAL");
                    } else {
                        MINVAL = null;
                    }

                    C_DOCPDK = Integer.parseInt(pdkProductObj.getC_DOCPDK());
                    System.out.println("ДокНорм" + C_DOCPDK);
                    C_METHODDEFINE = pdkProductObj.getC_METHODDEFINE();
                    System.out.println(C_METHODDEFINE);
                    DOCNORMATIVENAME = pdkProductObj.getDOCNORMATIVENAME();
                    System.out.println("НдМетод" + DOCNORMATIVENAME);
                    C_LABORATORY = pdkProductObj.getC_LABORATORY();
                    System.out.println("лаборатория" + C_LABORATORY);
                    COUNTRESEARCH = pdkProductObj.getCOUNTRESEARCH();
                    System.out.println(COUNTRESEARCH);
                    NOTE = pdkProductObj.getNOTE();
                    System.out.println(NOTE);
                    C_TYPETEST = pdkProductObj.getC_TYPETEST();
                    System.out.println(C_TYPETEST);

                    if (NORMTXT.equals("не допускается")) {
                        VALTXTOK = "Не обнаружено";
                        VALTXTBAD = "Обнаружено";
                    } else {
                    VALTXTOK = pdkProductObj.getVALTXTOK();
                    VALTXTBAD = pdkProductObj.getVALTXTBAD();
                    }
                    System.out.println("VALTXTOK" + VALTXTOK);
                    System.out.println("VALTXTBAD" + VALTXTBAD);
                    C_STANDART = pdkProductObj.getC_STANDARD();
                    System.out.println(C_STANDART);
                    Exp = pdkProductObj.getEXPONENT();
                }
                catch(NullPointerException e)
                {System.out.println("Пустая ячейка");}

                list_ID.add(ID);
                list_C_DOCPDK.add(C_DOCPDK);
                list_C_PRODUCT.add(C_PRODUCT);
                list_C_NOTE.add(C_NOTE);
                list_C_PACKING.add(C_PACKING);
                list_C_INGREDIENT.add(C_INGREDIENT);
                list_C_UNIT.add(C_UNIT);

                if (MAXVAL == null){
                    list_MAXVAL.add(null);
                }
                else{
                    list_MAXVAL.add(MAXVAL);
                }


                if (MINVAL == null){
                    list_MINVAL.add(null);
                }
                else{
                    list_MINVAL.add(MINVAL);
                }

                list_C_METHODDEFINE.add(C_METHODDEFINE);

                if (DOCNORMATIVENAME.equals("null")){
                    list_DOCNORMATIVENAME.add("null");
                }
                else {
                    list_DOCNORMATIVENAME.add("'" + DOCNORMATIVENAME + "'");
                }

                list_C_LABORATORY.add(C_LABORATORY);
                list_COUNTRESEARCH.add(COUNTRESEARCH);

                if (NOTE == null){
                    list_NOTE.add("null");
                }
                else {
                    list_NOTE.add("'" + NOTE + "'");
                }
                list_C_TYPETEST.add(C_TYPETEST);

                if (NORMTXT.equals("null")){
                    list_NORMTXT.add("null");
                }
                else {
                    list_NORMTXT.add("'" + NORMTXT + "'");
                    System.out.println("'" + NORMTXT + "'");
                }

                if (VALTXTOK.equals("null")){
                    list_VALTXTOK.add("null");
                }
                else {
                    list_VALTXTOK.add("'" + VALTXTOK + "'");
                }

                if (VALTXTBAD.equals("null")){
                    list_VALTXTBAD.add("null");
                }
                else {
                    list_VALTXTBAD.add("'" + VALTXTBAD + "'");
                }

                list_C_STANDART.add(C_STANDART);
                list_Exp.add(Exp);

            }

            //for (int id : list_ID){System.out.println(id);}
            //for (Double id : list_MAXVAL){System.out.println(id);}
            //for (Double id : list_MINVAL){System.out.println(id);}


            String min_line = "0.0";
            String max_line = "0.0";
            String normtxt = null;

            for (int i = 0; i < list_ID.size(); i++){
                int id = list_ID.get(i);
                System.out.println(" id" + id);
                int c_product = list_C_PRODUCT.get(i);
                int c_note = list_C_NOTE.get(i);
                int c_paking = list_C_PACKING.get(i);
                int c_ingredient = list_C_INGREDIENT.get(i);
                int c_unit = list_C_UNIT.get(i);
                Double maxval = list_MAXVAL.get(i);
                System.out.println("maxval " + maxval);
                Double minval = list_MINVAL.get(i);
                System.out.println("minval " + minval);
                int c_dokpdk = list_C_DOCPDK.get(i);
                int c_methoddefine = list_C_METHODDEFINE.get(i);
                String docnormativname = list_DOCNORMATIVENAME.get(i);
                int c_laboratory = list_C_LABORATORY.get(i);
                int countresearch = list_COUNTRESEARCH.get(i);
                String note = list_NOTE.get(i);
                int c_typetest = list_C_TYPETEST.get(i);
                String exp = list_Exp.get(i);


                if (!String.valueOf(minval).equals("null")){
                    min_line = String.valueOf(minval);
                    if (!exp.equals("")) {
                        int e = Integer.parseInt(exp);
                        double m = minval;
                        for (int j = 0; j < e; j++) {
                            m = m / 10;
                        }
                        min_line = String.valueOf(m);
                    }
                }
                if (!String.valueOf(maxval).equals("null")){
                    max_line = String.valueOf(maxval);
                    if (!exp.equals("")) {
                        int e = Integer.parseInt(exp);
                        double m = maxval;
                        for (int j = 0; j < e; j++) {
                            m = m / 10;
                        }
                        max_line = String.valueOf(m);
                    }
                }
                System.out.println(min_line);
                System.out.println(max_line);

                min_line = min_line.replaceAll("\\.",",");
                System.out.println(min_line);
                ArrayList<String> num_min = new ArrayList<>();
                for (String retval : min_line.split(",")) {
                    num_min.add(retval);
                }
                String min = num_min.get(0);
                System.out.println(min);

                String d_min = num_min.get(1);
                System.out.println(d_min);


                max_line = max_line.replaceAll("\\.",",");
                System.out.println(max_line);
                ArrayList<String> num_max = new ArrayList<>();
                for (String retval : max_line.split(",")) {
                    num_max.add(retval);
                }
                String max = num_max.get(0);
                System.out.println(max);


                String d_max = num_max.get(1);
                System.out.println(d_max);

                System.out.println("list_NORMTXT = " + list_NORMTXT.get(i));


                System.out.println(i + "exp = " + exp);
                if (list_NORMTXT.get(i).equals("'не допускается'")){
                    normtxt = list_NORMTXT.get(i);
                }
                else if(!exp.equals("")){
                        if (list_MAXVAL.get(i) == null && list_MINVAL.get(i) != null) {
                            if (d_min.equals("0")) {
                                normtxt = "'не менее " + min + "*10^" + exp + "'";
                            } else {
                                normtxt = "'не менее " + min + "," + d_min + "*10^" + exp + "'";
                            }
                        }

                        if (list_MAXVAL.get(i) != null && list_MINVAL.get(i) == null) {
                            if (d_max.equals("0")) {
                                normtxt = "'не более " + max + "*10^" + exp + "'";
                            } else {
                                normtxt = "'не более " + max + "," + d_max + "*10^" + exp + "'";
                            }
                        }
                        if (list_MAXVAL.get(i) != null && list_MINVAL.get(i) != null) {
                            if (d_min.equals("0") && d_max.equals("0")) {
                                normtxt = "'от " + min + "*10^" + exp + " до " + max + "*10^" + exp + "'";
                            }
                            if (!d_min.equals("0") && d_max.equals("0")) {
                                normtxt = "'от " + min + "*10^" + exp + "," + d_min + " до " + max + "*10^" + exp + "'";
                            }
                            if (d_min.equals("0") && !d_max.equals("0")) {
                                normtxt = "'от " + min + "*10^" + exp + " до " + max + "," + d_max + "*10^" + exp + "'";
                            }
                            if (!d_min.equals("0") && !d_max.equals("0")) {
                                normtxt = "'от " + min + "," + d_min + "*10^" + exp + " до " + max + "," + d_max + "*10^" + exp + "'";
                            }
                        }
                    }
                    else {
                        if (list_MAXVAL.get(i) == null && list_MINVAL.get(i) != null) {
                            if (d_min.equals("0")) {
                                normtxt = "'не менее " + min + "'";
                            } else {
                                normtxt = "'не менее " + min + "," + d_min + "'";
                            }
                        }

                        if (list_MAXVAL.get(i) != null && list_MINVAL.get(i) == null) {
                            if (d_max.equals("0")) {
                                normtxt = "'не более " + max + "'";
                            } else {
                                normtxt = "'не более " + max + "," + d_max + "'";
                            }
                        }
                        if (list_MAXVAL.get(i) != null && list_MINVAL.get(i) != null) {
                            if (d_min.equals("0") && d_max.equals("0")) {
                                normtxt = "'от " + min + " до " + max + "'";
                            }
                            if (!d_min.equals("0") && d_max.equals("0")) {
                                normtxt = "'от " + min + "," + d_min + " до " + max + "'";
                            }
                            if (d_min.equals("0") && !d_max.equals("0")) {
                                normtxt = "'от " + min + " до " + max + "," + d_max + "'";
                            }
                            if (!d_min.equals("0") && !d_max.equals("0")) {
                                normtxt = "'от " + min + "," + d_min + " до " + max + "," + d_max + "'";
                            }
                        }
                    }

                System.out.println(normtxt);

                String valtxtok = list_VALTXTOK.get(i);
                System.out.println(valtxtok);
                String valtxtbad = list_VALTXTBAD.get(i);
                System.out.println(valtxtbad);
                int c_standart = list_C_STANDART.get(i);

                int c_measuredtools = 0;
                int c_kindtest = 0;
                int c_damage = 0;
                int c_primaryproduct = 0;
                Double maxmaxval = null;
                int c_typeingredient = 0;
                String fault = "null";
                String minunit = "null";
                String danger = "null";
                String price = "null";
                String faultnd = "null";
                String faultndminus = "null";
                String percentfault = "null";
                String percentfaultnd = "null";
                String percentfaultndminus = "null";
                String exponent = "null";
                if(!exp.equals("")){
                    exponent = exp;
                }
                String decemalcount = "null";

                arrayList_SQL.add("INSERT INTO S_PDKPRODUCT (ID, C_NOTE, C_PACKING, C_UNIT, C_PRODUCT, C_TYPETEST, C_INGREDIENT, MINVAL, MAXVAL, C_METHODDEFINE, C_DOCPDK, NOTE, C_LABORATORY, COUNTRESEARCH," +
                        " NORMTXT, VALTXTOK, VALTXTBAD, DOCNORMATIVENAME, C_STANDARD, C_MEASUREDTOOLS, C_KINDTEST, C_DAMAGE, C_PRIMARYPRODUCT, MAXMAXVAL, C_TYPEINGREDIENT, FAULT, MINUNIT, DANGER,PRICE, FAULTND, " +
                        "PERCENTFAULT, FAULTNDMINUS, PERCENTFAULTND, PERCENTFAULTNDMINUS, EXPONENT, DECIMALCOUNT)\n" +
                        " VALUES (" + id + "," + c_note + "," + c_paking + "," + c_unit + "," + c_product + "," + c_typetest + "," + c_ingredient + "," +
                        minval + "," + maxval + "," + c_methoddefine + "," + c_dokpdk + "," + note + "," + c_laboratory + "," + countresearch + "," + normtxt + "," + valtxtok + "," + valtxtbad + "," + docnormativname + "," +
                        c_standart + ","  + c_measuredtools + "," + c_kindtest +  "," + c_damage +  "," + c_primaryproduct +  "," + maxmaxval +  "," + c_typeingredient +  "," + fault  + "," + minunit +  "," + danger +  "," + price +  "," + faultnd +  "," +
                        percentfault + ","  + faultndminus + "," + percentfaultnd + "," + percentfaultndminus +  "," + exponent +  "," + decemalcount + ");\n\n");

                min_line = "0.0";
                max_line = "0.0";
                normtxt = null;
                //minval = null;
                //maxval = null;

            }

            try {
                //FileWriter nFile = new FileWriter("C:/Users/dbarshhevskijj/Desktop/JAVA_EXEL/file_S_ProductPDK.sql"); // Создание файла
                BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream("C:/JAVA_EXEL/PDKConverter/_S_PDKProduct.sql"), "Windows-1251"));

                for(int j = 0; j < arrayList_SQL.size(); j++)
                {
                    bw.write(arrayList_SQL.get(j));
                }

                bw.flush();
                bw.close();
            }catch(IOException e) {
                System.out.print("Exception: no write file.sql");}
        }
}

