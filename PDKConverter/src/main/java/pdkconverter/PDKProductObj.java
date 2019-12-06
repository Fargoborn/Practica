package pdkconverter;

import org.apache.poi.ss.usermodel.Row;

import java.sql.SQLException;

public class PDKProductObj {

    private long ID = 0;
    private int C_NOTE = 0;
    private int C_PACKING = 0;
    private int C_UNIT = 0;
    private int C_PRODUCT = 0;
    private int C_TYPETEST = 40000;
    private int C_INGREDIENT = 0;
    private String MINVAL = "null";
    private String MAXVAL = "null";
    private int C_METHODDEFINE = 0;
    private String C_DOCPDK = null;
    private String NOTE = null;
    private int C_LABORATORY = 0;
    private int COUNTRESEARCH = 1;
    private String NORMTXT = "null";
    private String VALTXTOK = "null";
    private String VALTXTBAD = "null";
    private String DOCNORMATIVENAME = "null";
    private int C_STANDARD = 0;
    private int C_MEASUREDTOOLS = 0;
    private int C_KINDTEST = 0;
    private  int C_DAMAGE = 0;
    private int C_PRIMARYPRODUCT = 0;
    private String MAXMAXVAL = "null";
    private int C_TYPEINGREDIENT = 0;
    private String FAULT = "null";
    private String MINUNIT = "null";
    private String DANGER = "null";
    private String PRICE = "null";
    private String FAULTND = "null";
    private String PERCENTFAULT = "null";
    private String FAULTNDMINUS = "null";
    private String PERCENTFAULTND = "null";
    private String PERCENTFAULTNDMINUS = "null";
    private String EXPONENT = "null";
    private String DECIMALCOUNT = "null";

    public PDKProductObj() {
    }

    public PDKProductObj(long ID, int c_PRODUCT, int c_NOTE, int c_INGREDIENT, int c_UNIT, String MINVAL, String MAXVAL, String c_DOCPDK, int c_LABORATORY, String EXPONENT) {
        this.ID = ID;
        this.C_NOTE = c_NOTE;
        this.C_UNIT = c_UNIT;
        this.C_PRODUCT = c_PRODUCT;
        this.C_INGREDIENT = c_INGREDIENT;
        this.MINVAL = MINVAL;
        this.MAXVAL = MAXVAL;
        this.C_DOCPDK = c_DOCPDK;
        this.C_LABORATORY = c_LABORATORY;
        this.EXPONENT = EXPONENT;
    }

    public long getID() {
        return ID;
    }

    public int getC_NOTE() {
        return C_NOTE;
    }

    public int getC_PACKING() {
        return C_PACKING;
    }

    public int getC_UNIT() {
        return C_UNIT;
    }

    public int getC_PRODUCT() {
        return C_PRODUCT;
    }

    public int getC_TYPETEST() {
        return C_TYPETEST;
    }

    public int getC_INGREDIENT() {
        return C_INGREDIENT;
    }

    public String getMINVAL() {
        return MINVAL;
    }

    public String getMAXVAL() {
        return MAXVAL;
    }

    public String getC_DOCPDK() {
        return C_DOCPDK;
    }

    public int getC_METHODDEFINE() {
        return C_METHODDEFINE;
    }

    public String getNOTE() {
        return NOTE;
    }

    public int getC_LABORATORY() {
        return C_LABORATORY;
    }

    public int getCOUNTRESEARCH() {
        return COUNTRESEARCH;
    }

    public String getNORMTXT() {
        return NORMTXT;
    }

    public String getVALTXTOK() {
        return VALTXTOK;
    }

    public String getVALTXTBAD() {
        return VALTXTBAD;
    }

    public String getDOCNORMATIVENAME() {
        return DOCNORMATIVENAME;
    }

    public int getC_STANDARD() {
        return C_STANDARD;
    }

    public int getC_MEASUREDTOOLS() {
        return C_MEASUREDTOOLS;
    }

    public int getC_KINDTEST() {
        return C_KINDTEST;
    }

    public int getC_DAMAGE() {
        return C_DAMAGE;
    }

    public int getC_PRIMARYPRODUCT() {
        return C_PRIMARYPRODUCT;
    }

    public String getMAXMAXVAL() {
        return MAXMAXVAL;
    }

    public int getC_TYPEINGREDIENT() {
        return C_TYPEINGREDIENT;
    }

    public String getFAULT() {
        return FAULT;
    }

    public String getMINUNIT() {
        return MINUNIT;
    }

    public String getDANGER() {
        return DANGER;
    }

    public String getPRICE() {
        return PRICE;
    }

    public String getFAULTND() {
        return FAULTND;
    }

    public String getPERCENTFAULT() {
        return PERCENTFAULT;
    }

    public String getFAULTNDMINUS() {
        return FAULTNDMINUS;
    }

    public String getPERCENTFAULTND() {
        return PERCENTFAULTND;
    }

    public String getPERCENTFAULTNDMINUS() {
        return PERCENTFAULTNDMINUS;
    }

    public String getEXPONENT() {
        return EXPONENT;
    }

    public String getDECIMALCOUNT() {
        return DECIMALCOUNT;
    }
}
