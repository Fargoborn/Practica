package main.java.exelerator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Iterator;

/**
 * @version $Id$
 * @since 0.1
 */
public class Main {

    public static void main(String[] args) throws IOException {

        System.out.println("Что обрабатываем?");
        System.out.println("1. C_Docnorm");
        System.out.println("2. C_Ingr");
        System.out.println("3. C_Prod");
        System.out.println("*****************");

        InputStreamReader isr = new InputStreamReader(System.in);
        BufferedReader reader  = new BufferedReader(isr);
        String line = reader.readLine();

        if(line.equals("1")) {
        C_DocNorm c_docNorm = new C_DocNorm();
        c_docNorm.run_C_DocNorm();
        }

        if(line.equals("2")) {
            C_Ingredient c_ingredient = new C_Ingredient();
            c_ingredient.run_C_Ingredient();
        }
    }
}
