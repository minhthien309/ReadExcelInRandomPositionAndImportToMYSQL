
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.nio.charset.Charset;
import java.util.Properties;
import javax.swing.JOptionPane;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author Thien
 */

public class EditMe {
    static FileInputStream inputStream;
    
    public static String FILE_INPUT_STREAM = "";
    public static String[] columnArray = {};
    public static String tableName = ""; //Ten table
    public static String keyName = ""; //Ten khoa chinh
    public static int keyCol = 0; //Vi tri khoa chinh
    public static int sheet =  0; //Sheet 0 ứng với khách hàng bồn, 1 ứng với khách hàng phi

    public static String checkName = ""; //Ten kiem tra trung
    public static int checkCol = 2; //Vi tri kiem tra trung

    public static String firstCol = ""; //Ten cot dau tien
    public static String lastCol = ""; //Ten cot cuoi cung

        //Phan ket noi csdl
    public static String hostName = ""; //Dia chi host
    public static String dbName = ""; //Ten Database
    public static String dbUser = ""; //Ten User
    public static String dbPass = ""; //Pass User
    public static int check = 0;
    
    public static void getPropValues() throws IOException {

        try {
            Properties prop = new Properties();
            File  file = new File (EditMe.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath());
            String path = file.getParentFile().getPath();
            String propFileName = path+"\\config.txt";
            //inputStream = EditMe.class.getClassLoader().getResourceAsStream(propFileName);
            inputStream = new FileInputStream(propFileName);
                    
            if (inputStream != null) {
            prop.load(new InputStreamReader(inputStream,Charset.forName("UTF-8")));
            check = 1;
            } else {
                prop.setProperty("FILE_INPUT_STREAM","excel.xlsx");
                prop.setProperty("tableName","db_tableName");
                prop.setProperty("Columns","colum1, column2");
                prop.setProperty("keyName","");
                prop.setProperty("keyCol","0");
                prop.setProperty("sheet","0");
                prop.setProperty("checkKey","");
                prop.setProperty("checkCol","0");
                prop.setProperty("cotDau","");
                prop.setProperty("cotCuoi","");
                prop.setProperty("hostName","localhost");
                prop.setProperty("dbName","");
                prop.setProperty("dbUser","");
                prop.setProperty("dbPass","");
                prop.store(new OutputStreamWriter( new FileOutputStream(path+"\\config.txt"), "UTF-8"),null);

                throw new FileNotFoundException("Khong tim thay file config. Chuong trinh se tao file mau");
            }


	// get the property value and print it out
        FILE_INPUT_STREAM = prop.getProperty("FILE_INPUT_STREAM");
        tableName = prop.getProperty("tableName");
        
        String[] prop_column = prop.getProperty("Columns").split(",");
        columnArray = new String[prop_column.length];
        //columnArray = prop_column;
        for (int i=0;i<prop_column.length;i++){
            columnArray[i] = prop_column[i];
        }
        
        keyName = prop.getProperty("keyName");
        keyCol = Integer.parseInt(prop.getProperty("keyCol"));
        sheet = Integer.parseInt(prop.getProperty("sheet"));
        
        checkName = prop.getProperty("checkKey");
        checkCol = Integer.parseInt(prop.getProperty("checkCol"));
        
        firstCol = prop.getProperty("cotDau");
        lastCol = prop.getProperty("cotCuoi");
        hostName = prop.getProperty("hostName");
        dbName = prop.getProperty("dbName");
        dbUser = prop.getProperty("dbUser");
        dbPass = prop.getProperty("dbPass");
        
        } catch (Exception e) {
            //System.out.println("Exception: " + e);
            JOptionPane.showMessageDialog(null, "Không tìm thấy file config. File config sẽ được tạo mới!");
            System.exit(0);
        } finally {
            inputStream.close();
        }
    }

    
}

