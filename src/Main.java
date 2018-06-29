
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
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
public class Main {
    
    static SplashScreen splashscreen = new SplashScreen();
    
    public static void reader() throws FileNotFoundException, IOException, InvalidFormatException, ClassNotFoundException, SQLException {
        if(!new File(EditMe.FILE_INPUT_STREAM).isFile()){
            JOptionPane.showMessageDialog(null, "Không tìm thấy File Excel");
            System.exit(0);
        }
        splashscreen.setVisible(true);
        Workbook workbook = WorkbookFactory.create(new FileInputStream(EditMe.FILE_INPUT_STREAM));
        Sheet sheet = workbook.getSheetAt(EditMe.sheet);
        int totalRows = sheet.getLastRowNum();
        ArrayList<Point> listTable = new ArrayList<>();
        ArrayList<Point> listMaxY = new ArrayList<>();
        int minColIx;
        int maxColIx;

        DataFormatter fmt = new DataFormatter();
        for (int i = 0; i < totalRows; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                minColIx = row.getFirstCellNum();
                maxColIx = row.getLastCellNum();
                for (int colIx = minColIx; colIx < maxColIx; colIx++) {
                    String check = fmt.formatCellValue(row.getCell(colIx));
                    if (check != null && (check.equals(EditMe.firstCol) || check.equals(EditMe.lastCol))) {
                        Point start = new Point();
                        Cell cell = row.getCell(colIx);
                        start.setXY(cell.getColumnIndex(), i);
                        start.setNote(cell.getStringCellValue());
                        listTable.add(start);
                    }
                }
            }
        }

        for (int i = 0; i < listTable.size(); i++) {
            if (listTable.get(i).getNote().equals(EditMe.firstCol)) {
                for (int j = listTable.get(i).getY(); j < totalRows + 2; j++) {
                    Row row = sheet.getRow(j);
                    try {
                        Cell cell = row.getCell(listTable.get(i).getX());
                        if (cell == null) {
                            Point maxY = new Point();
                            Row rowBack = sheet.getRow(j - 1);
                            Cell cellUp = rowBack.getCell(listTable.get(i).getX());
                            maxY.setXY(cellUp.getColumnIndex(), j - 1);
                            maxY.setNote("MaxY");
                            listMaxY.add(maxY);
                            break;
                        }
                    } catch (Exception e) {
                        Point maxY = new Point();
                        Row rowBack = sheet.getRow(j - 2);
                        Cell cellUp = rowBack.getCell(listTable.get(i).getX());
                        maxY.setXY(cellUp.getColumnIndex(), j - 1);
                        maxY.setNote("MaxY");
                        listMaxY.add(maxY);
                        break;
                    }
                }
            }
        }

        int x = 0;
        int y = 0;
        int Ox = 0;
        int Oy = 0;
        int num = 1;
        int z = 0;
        for (int i = 0; i < listTable.size(); i = i + 2) {

            x = listTable.get(i).getX();
            y = listTable.get(i).getY();
            Ox = listTable.get(i + 1).getX();
            Oy = listMaxY.get(num - 1).getY();

            for (int oy = y; oy < Oy + 1; oy++) {
                z++;
                Row r = sheet.getRow(oy);
                String data[] = new String[EditMe.columnArray.length];
                int index = 0;
                for (int ox = x; ox < Ox + 1; ox++) {
                    Cell cellRowNext = r.getCell(ox);
                    if (z != 1) {
                        //if (index != 0) {
                        String cellData = fmt.formatCellValue(cellRowNext);
                        cellData = cellData.replace("'","''");
                        data[index] = "'" + cellData + "'"; //Get data from excel
                        //}
                        index++;
                    }
                }
                Sql(data);
            }
            z = 0;
            num++;
        }

    }

    public static void Sql(String[] data) throws ClassNotFoundException, SQLException {
        Connection connection = null;
        try {
            connection = SqlConnector.getMySQLConnection();
            if(connection.isClosed()){
                splashscreen.setVisible(false);
                JOptionPane.showMessageDialog(null, "Có lỗi xảy ra trong quá trình kết nối cơ sở dữ liệu, vui lòng kiểm tra lại file config!");
                System.exit(0);
            }
            Statement statement = connection.createStatement();
            //ResultSet rs = null;

            String sqlInsert = "INSERT INTO `"+EditMe.tableName+"`";
            String columnNames = "";
            String valueString = "";
            String value[] = data;
            int columnCount = value.length;// get value from excel sheet

            for (int i = 0; i < columnCount; i++) {
                if (i != 0) {
                    columnNames += ",";
                    valueString += ",";
                }
                    columnNames += "`" + EditMe.columnArray[i] + "`";
                    valueString += value[i];
                    
                    
                
            }
            
            sqlInsert += "(" + columnNames + ") VALUES (" + valueString + ");";

            sqlInsert = sqlInsert.replace("[", "");
            sqlInsert = sqlInsert.replace("]", "");
            String keyName = EditMe.keyName;
            System.out.println(sqlInsert);
            if(!keyName.equals("")){
                String checkQuery = "SELECT `"+EditMe.checkName+"` AS countValue FROM `"+EditMe.tableName+"` WHERE `"+EditMe.keyName+"` = " + value[EditMe.keyCol] + "";
                statement.execute(checkQuery);
                ResultSet resultSet = statement.getResultSet();
                if (resultSet.next() == true) {
                    int Amounts = resultSet.getInt("countValue") + Integer.parseInt(value[EditMe.checkCol].replace("'", ""));
                    String updateQuery = "UPDATE `"+EditMe.tableName+"` SET `"+EditMe.checkName+"`=" + String.valueOf(Amounts) + " WHERE `"+EditMe.keyName+"`=" + value[EditMe.keyCol] + ";";
                    statement.executeUpdate(updateQuery);
                } else {
                    statement.executeUpdate(sqlInsert);

                }
            }
            else{
                    statement.executeUpdate(sqlInsert);
////                   //statement.executeUpdate("INSERT INTO `tbl_receipt_delivery_place`(`place_type`,`name`,`contact_note`,`warehouse_note`,`address`) VALUES ('1','AKA VINA','0909 285 046','CÓ XE NÂNG,5H30 NGHỈ LÀM','');");
            }
            
        } catch (SQLException ex) {
//            splashscreen.setVisible(false);
//            JOptionPane.showMessageDialog(null, "Có lỗi xảy ra trong quá trình kết nối cơ sở dữ liệu, vui lòng kiểm tra lại file config!");
            Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
//            System.exit(0);
        }
        finally{
            connection.close();
        }
    }

    public static void main(String args[]) throws IOException, FileNotFoundException, InvalidFormatException, SQLException, ClassNotFoundException {
        //EditMe editme = new EditMe();
        EditMe.getPropValues();
        if(EditMe.check==1){
            reader();
            splashscreen.setVisible(false);
            JOptionPane.showMessageDialog(null, "Đã xong!");
            System.exit(0);
        }
        
    }
}
