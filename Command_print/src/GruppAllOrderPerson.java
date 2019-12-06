import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

public class GruppAllOrderPerson {
    String allOrder = "";
    String fio = "";
    String order = "";

    public HashMap<String, String> gruppAllOrderPerson(ArrayList<String> _order, ArrayList<String> _fio) {
        HashMap<String, String> hashMap = new HashMap();
        ArrayList<String> l_order = new ArrayList<>();
        hashMap.put("a", "s");
        for (int i = 0; i < _fio.size(); i++){
            fio = String.valueOf(_fio.get(i)).trim();
            order = String.valueOf(_order.get(i)).trim();
            for (int j = 1; j < _fio.size(); j++){
                l_order.add("");
            if (fio.equals(String.valueOf(_fio.get(j)).trim())) {
                l_order.set(i, l_order.get(i) + " ," + String.valueOf(_order.get(j)).trim());
            }
        }}
        for (int h = 0; h < _fio.size(); h++) {
            //System.out.println("ФИО =  " + _fio.get(h).trim() + " № приказа = " + l_order.get(h));
            hashMap.put(String.valueOf(_fio.get(h)).trim(), l_order.get(h));
            }
        System.out.println("******************************************");
        return hashMap;
    }
}

