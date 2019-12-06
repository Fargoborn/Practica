package pdkconverter;

public class PDKObj {

    private Integer ID;
    private Integer ingredient;
    private Integer v_probe;
    private Integer C_Unit;
    private double pdk;
    private String note_pdk;
    private Integer class_dang;
    private Integer damage;
    private Integer ND;
    private Integer C_Laboratory;
    private Integer C_Typetest;




    PDKObj (Integer ID, Integer ingredient, Integer v_probe, Integer C_Unit, double pdk, String note_pdk, Integer class_dang, Integer damage, Integer ND, Integer C_Laboratory,
             Integer C_Typetest){
        this.ID = ID;
        this.ingredient = ingredient;
        this.v_probe = v_probe;
        this.C_Unit = C_Unit;
        this.pdk = pdk;
        this.note_pdk = note_pdk;
        this.class_dang = class_dang;
        this.damage = damage;
        this.ND = ND;
        this.C_Laboratory = C_Laboratory;
        this.C_Typetest = C_Typetest;

    }


    public Integer getID() {
        return ID;
    }

    public Integer getingredient() {
        return ingredient;
    }


    public Integer getC_Unit() {
        return C_Unit;
    }

    public Integer getV_probe() {
        return v_probe;
    }

    public Double getpdk() {
        return pdk;
    }

    public String getnote_pdk() {
        return note_pdk;
    }

    public Integer getclass_dang() {
        return class_dang;
    }

    public Integer getdamage() {
        return damage;
    }

    public Integer getND() {
        return ND;
    }

    public Integer getC_Laboratory() {
        return C_Laboratory;
    }

    public Integer getC_Typetest() {
        return C_Typetest;
    }
}