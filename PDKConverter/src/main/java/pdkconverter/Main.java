package pdkconverter;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.sql.*;


/**
 * @version $Id$
 * @since 0.1
 */
public class Main {
    public static void main(String[] args) throws IOException, SQLException, ClassNotFoundException {

        //AdditionUI additionUI = new AdditionUI();
        //additionUI.setVisible(true);

        ConverterUI converterUI = new ConverterUI();
        converterUI.setVisible(true);
    }
}

