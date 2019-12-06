package padeg_exel;

import java.util.ArrayList;

public class  Get_NativePadeg {
//*
    String fio_NativePadeg;

    public Get_NativePadeg(String fio_NativePadeg) {
        this.fio_NativePadeg = fio_NativePadeg;
    }

    public String getGet_NativePadeg() {

    //Get_Person g_p = new Get_Person(new ArrayList<Person>());
    //ArrayList list_Person = null;

    //{
        //try {
           //list_Person = new ArrayList(g_p.Get_Person());
       // } catch (FileNotFoundException e) {
          //  e.printStackTrace();
        //}
    //}

    String FEnds; //Окончание фамилий в родительном падеже
    String FNativEnds; //Окончание фамилий в именительном падеже
    String NFEnds; //Окончание имён в родительном падеже
    String NFNativEnds; //Окончание имён в именительном падеже
    String NFF_Ends; //Окончание имён в родительном падеже
    String NFF_NativEnds; //Окончание имён в именительном падеже
    // 'Окончания родительного падежа для фамилий женского рода
    String FEMALECASEENDS = "ую он о ову у д х н к т и ";
    // 'Окончания именительного падежа для фамилий женского рода
    String FEMALEENDS = "ая он о ова а д х н к т и ";

    //'Окончания родительного падежа для фамилий мужского рода
    String MALECASEENDS = "ую у он о ова я ого а х н к т и ";
    //'Окончания именительного падежа для фамилий мужского рода.
    //'Знак $$ означает, что окончания нет
    String MALEENDS = "ая а он о ов ь ий $$ х н к т и";

    // 'Окончания родительного падежа для имён женского рода
    String N_FEMALECASEENDS = "у ю ик ь";
    // 'Окончания родительного падежа для имён женского рода
    String N_FEMALEENDS = "а я ик ь";

    // 'Окончания родительного падежа для имён мужского рода
    String N_MALECASEENDS = "ю у я а ла ила";
    // 'Окончания родительного падежа для имён мужского рода
    String N_MALEENDS = "я а й $$ ел ил";

    // 'Окончания родительного падежа для отчества женского рода
    String NF_FEMALECASEENDS = "у $$";
    // 'Окончания родительного падежа для отчества женского рода
    String NF_FEMALEENDS = "а $$";

    // 'Окончания родительного падежа для отчества мужского рода
    String NF_MALECASEENDS = "а $$";
    // 'Окончания родительного падежа для отчества мужского рода
    String NF_MALEENDS = "$$ $$";

    //for (int i = 0; i < list_Person.size(); i++){
    //System.out.println("Номер по приказу  " + i);
    //}

    ArrayList<String> list_f = new ArrayList<>();
    ArrayList<String> list_i = new ArrayList<>();
    ArrayList<String> list_o = new ArrayList<>();
    ArrayList<String> list_fio = new ArrayList<>();

    String L_f = "";
    String L_i = "";
    String L_o = "";

    //Iterator<Person> personIterator = list_Person.iterator();

        //if (personIterator.hasNext()) {
        //while (personIterator.hasNext()) {
            //Person currPerson = personIterator.next();
            //if (currPerson.get$name() != null){
                //System.out.println("Печать ФИО  " + currPerson.get$name());

                //String[] FIO = currPerson.get$name().split(" ");

                String[] FIO = fio_NativePadeg .split(" ");

                //System.out.println("Печать падежей ++++++++++++++++++++++++++++++++++++++++++++++++++");

                PadegBean pb = new PadegBean();

                pb.setFirstName(FIO[1]);

                if(FIO.length < 3){pb.setMiddleName("$$");}
                else pb.setMiddleName(FIO[2]);
                pb.setLastName(FIO[0]);
                pb.setSexStr("auto");
                //System.out.print(pb.getLastName() + " " + pb.getFirstName() + " " + pb.getMiddleName());

                //for (int i = 0; i < 6; i++){
                //System.out.println(pb.getResultItems().get(i).getFio());
                //}

                //System.out.println("Печать пола ++++++++++++++++++++++++++++++++++++++++++++++++++");
                String SexStr = "человек";
                if (pb.getResultItems().get(1).getSex() == 0){SexStr = "женский";}
                if (pb.getResultItems().get(1).getSex() == 1){SexStr = "мужской";}

                //System.out.println("ФИО - " + currPerson.get$name() + " пол: " + SexStr);

                if (SexStr.equals("женский")){
                    FEnds = FEMALECASEENDS;
                    FNativEnds = FEMALEENDS;
                    NFEnds = N_FEMALECASEENDS;
                    NFNativEnds = N_FEMALEENDS;
                    NFF_Ends = NF_FEMALECASEENDS;
                    NFF_NativEnds = NF_FEMALEENDS;}
                else{
                    FEnds = MALECASEENDS;
                    FNativEnds = MALEENDS;
                    NFEnds = N_MALECASEENDS;
                    NFNativEnds = N_MALEENDS;
                    NFF_Ends = NF_MALECASEENDS;
                    NFF_NativEnds = NF_MALEENDS;}

               // System.out.print(pb.getLastName() + " " + pb.getFirstName() + " " + pb.getMiddleName());

                String arEnds [] = FEnds.split(" ");
                String Nativ_arEnds [] = FNativEnds.split(" ");

                String arNEnds [] = NFEnds.split(" ");
                String Nativ_arNEnds [] = NFNativEnds.split(" ");

                String arNFEnds [] = NFF_Ends.split(" ");
                String Nativ_arNFEnds [] = NFF_NativEnds.split(" ");

                //бработка Фамилии
                for (int i = 0; i < arEnds.length; i++){
                    //System.out.println(pb.getLastName() + "111");
                    //System.out.println("окончание: " + pb.getLastName().substring(pb.getLastName().length() - arEnds[i].length()) + " " + arEnds[i]);
                    if(pb.getLastName().substring(pb.getLastName().length() - arEnds[i].length()).equals(arEnds[i])){
                        L_f = pb.getLastName().substring(0, pb.getLastName().length() - arEnds[i].length()) + Nativ_arEnds[i];
                        L_f = L_f.replace("$$", "");
                        //System.out.println("Именительный падеж: " + L_f + "222" + pb.getLastName());
                    }
                }
                list_f.add(L_f);

                //бработка Имени
                for (int i = 0; i < arNEnds.length; i++){
                    if(pb.getFirstName().substring(pb.getFirstName().length() - arNEnds[i].length()).equals(arNEnds[i])){
                        L_i = pb.getFirstName().substring(0, pb.getFirstName().length() - arNEnds[i].length()) + Nativ_arNEnds[i];
                        L_i = L_i.replace("$$", "");
                        //System.out.println("Именительный падеж: " + L_i);
                    }
                }
                list_i.add(L_i);

                //бработка Отчества
                for (int i = 0; i < arNFEnds.length; i++){
                    if(pb.getMiddleName().substring(pb.getMiddleName().length() - arNFEnds[i].length()).equals(arNFEnds[i])){
                        L_o = pb.getMiddleName().substring(0, pb.getMiddleName().length() - arNFEnds[i].length()) + Nativ_arNFEnds[i];
                        L_o = L_o.replace("$$", "");
                        //System.out.println("Именительный падеж: " + L_o);
                    }
                }
                list_o.add(L_o);

        //Сборка ФИО
        for (int i = 0; i < list_f.size(); i++){
        list_fio.add(list_f.get(i) + " " + list_i.get(i) + " " + list_o.get(i));
        //System.out.println("Именительный падеж: " + list_f.get(i) + " " + list_i.get(i) + " " + list_o.get(i));
        fio_NativePadeg = list_f.get(i) + " " + list_i.get(i) + " " + list_o.get(i);
    }
        return fio_NativePadeg;
    }
}
