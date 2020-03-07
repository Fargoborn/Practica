package setrandomtime;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class Mine {
    public static void main(String[] args)throws Exception {

      BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
        System.out.println("Укажите рабочий каталог:");
        String _path = reader.readLine();

        System.out.println("Введите дату в формате ddmmyy:");
        String _date = reader.readLine();
        SimpleDateFormat dateFormat = new SimpleDateFormat("ddMMyy");
        System.out.println(dateFormat.parse(_date));
        Date sdate = dateFormat.parse(_date);

        System.out.println("Введите время начала работы (часы без минут)");
        int sw_time = Integer.parseInt(reader.readLine());
        System.out.println(sw_time);
        long nsw_time = (sw_time * 3600000);
        System.out.println(new Date(sdate.getTime() + nsw_time));

        System.out.println("Введите время конца работы (часы без минут)");
        int fw_time = Integer.parseInt(reader.readLine());
        System.out.println(fw_time);
        long nfw_time = (fw_time * 3600000);
        System.out.println(new Date(sdate.getTime() + nfw_time));

        System.out.println("Введите требуемый промежуток в минутах");
        long time_cr = Integer.parseInt(reader.readLine());
        System.out.println(time_cr);
        time_cr = (time_cr * 60000);


        System.out.println("*_____________________________*");
        System.out.println("Запущена обработка файлов");

        File dir = new File(_path);
        File[] arrFiles = dir.listFiles();
        List<File> list = Arrays.asList(arrFiles);
        System.out.println(list.size());

        //long time_cr = (nfw_time - nsw_time) / list.size();
        //System.out.println(time_cr);
        long j = time_cr;


        long modDate = 0;
        for (int i = 0; i < list.size(); i++){
            File file = new File(String.valueOf(list.get(i)));
                if(file.exists()){
                    System.out.println(file);
                    long s_t = sdate.getTime() + nsw_time;
                    modDate = s_t + time_cr;
                    boolean result = file.setLastModified(modDate);
                    //writer.write("touch -t " + modDate + " " + list.get(i) + " \n");
                    // check if the rename operation is success
                    if(result){
                        System.out.println("Operation Success");
                        System.out.println("lastModifiedTime_modDate: " + modDate);
                    }
                    time_cr = time_cr + j + ((long)( 0 +(Math.random() * 2.5)) * 60000);
                    System.out.println(time_cr/3600000);
                }
            if(file.isDirectory()){
                File[] listFiles = file.listFiles();
                if (listFiles != null) {
                    List<File> listt = Arrays.asList(listFiles);
                    Collections.sort(listt);
                    for (File file1 : listt) {
                        System.out.println(file1);
                        long s_t = sdate.getTime() + nsw_time;
                        modDate = s_t + time_cr;
                        boolean result = file1.setLastModified(modDate);
                        //writer.write("touch -t " + modDate + " " + list.get(i) + " \n");
                        // check if the rename operation is success
                        if(result){
                            System.out.println("Operation Success");
                            System.out.println("lastModifiedTime_modDate: " + modDate);
                        }
                        time_cr = time_cr + j + ((long)( 0 +(Math.random() * 2.5)) * 60000);
                        System.out.println(time_cr/3600000);
                    }
                }
            }
        }
  }
}
