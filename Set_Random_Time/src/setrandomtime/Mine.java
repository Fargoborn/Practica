package setrandomtime;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class Mine {
    public static void main(String[] args)throws Exception {

        long modDate = 0;

        System.out.println("Введите время начала работы (часы без минут)");
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
        int sw_time = Integer.parseInt(reader.readLine());
        System.out.println(sw_time);
        long nsw_time = (sw_time * 3600000) - 1800000;
        System.out.println(nsw_time);

        System.out.println("Введите время конца работы (часы без минут)");
        BufferedReader freader = new BufferedReader(new InputStreamReader(System.in));
        int fw_time = Integer.parseInt(freader.readLine());
        System.out.println(fw_time);
        long nfw_time = (fw_time * 3600000) - 3600000;
        System.out.println(nfw_time);

        System.out.println("*_____________________________*");
        System.out.println("Запущена обработка файлов");

        File dir = new File("C:\\JAVA_EXEL\\ORDER_EXEL\\");
        File[] arrFiles = dir.listFiles();
        List<File> list = Arrays.asList(arrFiles);
        System.out.println(list.size());

        long time_cr = (nfw_time - nsw_time) / list.size();
        System.out.println(time_cr);
        long j = time_cr;
        Date date = new Date();
        long d_t = (long) date.getTime();

        SimpleDateFormat format = new SimpleDateFormat("HH");
        int i_t = Integer.parseInt(format.format(date.getTime()));
        System.out.println(i_t);

        long s_t = d_t - i_t * 3600000 + nsw_time;

        //Files.walk(Paths.get("C:\\JAVA_EXEL\\TIME_CORR"))
                //.filter(Files::isRegularFile)
               // .map(Path::toFile)
               // .forEach(File::delete);

        //File file_tc = new File("C:\\JAVA_EXEL\\TIME_CORR\\time_corr.bat");
        //FileWriter writer = new FileWriter(file_tc);


        for (int i = 0; i < list.size(); i++){
            File file = new File(String.valueOf(list.get(i)));
        if(file.exists()){
            System.out.println(file);
            // initialize calendar object
            //Calendar cal_ = Calendar.getInstance();
            modDate = s_t + time_cr;
            boolean result = file.setLastModified(modDate);

            //writer.write("touch -t " + modDate + " " + list.get(i) + " \n");

            // check if the rename operation is success
            if(result){
                System.out.println("Operation Success");
                System.out.println("lastModifiedTime_modDate: " + modDate);
            }
        }
            time_cr = time_cr + j + ((long)( 0 +(Math.random() * 5)) * 60000);
            System.out.println(time_cr/3600000);
    }

        //writer.close();

  }
}
