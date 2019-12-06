package padeg_exel;

import java.util.Date;

public class Person {

    String $name;
    double $tabel;
    String $S_Unit;
    String $New_S_Unit;
    String $Present_pos;
    String $New_Present_pos;
    String $Prise;
    String $date_new_transfer;
    double $N_Ord;
    String $date_Ord;
    String $date_OrdOrd;

    Person (String $name, double $tabel, String $S_Unit , String $New_S_Unit, String $Present_pos, String $New_Present_pos, String $Prise, String $date_new_transfer, double $N_Ord, String $date_Ord, String $date_OrdOrd)
    {
        this.$name = $name;
        this.$tabel = $tabel;
        this.$S_Unit = $S_Unit;
        this.$New_S_Unit = $New_S_Unit;
        this.$Present_pos = $Present_pos;
        this.$New_Present_pos = $New_Present_pos;
        this.$Prise = $Prise;
        this.$date_new_transfer = $date_new_transfer;
        this.$N_Ord = $N_Ord;
        this.$date_Ord = $date_Ord;
        this.$date_OrdOrd = $date_OrdOrd;
    }

    Person (String $name){
        this.$name = $name;
    }

    public String get$name(){
        return $name;
    }

    public double get$tabel(){
        return $tabel;
    }

    public String get$S_Unit(){
        return $S_Unit;
    }

    public String get$New_S_Unit(){
        return $New_S_Unit;
    }

    public String get$Present_pos(){
        return $Present_pos;
    }

    public String get$New_Present_pos(){
        return $New_Present_pos;
    }

    public String get$Prise(){
        return $Prise;
    }

    public String get$date_new_transfer(){
        return $date_new_transfer;
    }

    public double get$N_Ord(){
        return $N_Ord;
    }

    public String get$date_Ord(){
        return $date_Ord;
    }

    public String $date_OrdOrd(){
        return $date_OrdOrd;
    }
}
