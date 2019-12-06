import java.util.Date;

public  class Person {
    int $N_Ord;
    String $date_Ord;
    int $tabel;
    String $name;
    String $S_Unit;
    String $Present_pos;
    String $adres_city;
    String $adres_in;
    String $bisnes;
    Date $date_S;
    String $date_s_d;
    String $date_s_m;
    String $date_s_y;
    Date $date_F;
    String $date_f_d;
    String $date_f_m;
    String $date_f_y;
    String $date_Ord_y;

    Person (){ }

    Person(int $N_Ord, String $date_Ord, String $date_Ord_y, int $tabel, String $name, String $S_Unit, String $Present_pos, String $adres_city, String $adres_in, String $bisnes, Date $date_S, String $date_s_d,
           String $date_s_m, String $date_s_y, Date $date_F, String $date_f_d, String $date_f_m, String $date_f_y)
    {
        this.$N_Ord = $N_Ord;
        this.$date_Ord = $date_Ord;
        this.$date_Ord_y = $date_Ord_y;
        this.$tabel = $tabel;
        this.$name = $name;
        this.$S_Unit = $S_Unit;
        this.$Present_pos = $Present_pos;
        this.$adres_city = $adres_city;
        this.$adres_in = $adres_in;
        this.$bisnes = $bisnes;
        this.$date_S = $date_S;
        this.$date_s_d = $date_s_d;
        this.$date_s_m = $date_s_m;
        this.$date_s_y = $date_s_y;
        this.$date_F = $date_F;
        this.$date_f_d = $date_f_d;
        this.$date_f_m = $date_f_m;
        this.$date_f_y = $date_f_y;
    }

    public String getPersonName() {
        return $name;
    }

    public int getPersonOrder() {
        return $N_Ord;
    }
}