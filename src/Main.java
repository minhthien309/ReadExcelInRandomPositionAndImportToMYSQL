
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

    private static String FILE_INPUT_STREAM = "TNK.xlsx";
    public static int count = 0;


    public static void reader() throws FileNotFoundException, IOException, InvalidFormatException, ClassNotFoundException {
        Workbook workbook = WorkbookFactory.create(new FileInputStream(FILE_INPUT_STREAM));

        Sheet sheet = workbook.getSheetAt(0);

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
                    if (check != null && (check.equals("STT") || check.equals("Phần mềm"))) {
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
            if (listTable.get(i).getNote().equals("STT")) {
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
                String data[] = new String[5];
                int index = 0;
                for (int ox = x; ox < Ox + 1; ox++) {
                    Cell cellRowNext = r.getCell(ox);
                    if (z != 1) {
                        if (index != 0) {
                        data[index] = "'" + fmt.formatCellValue(cellRowNext) + "'"; //Get data from excel
                        }
                        index++;
                    }
                }
                Sql(data);
            }
            z = 0;
            num++;
        }

    }

    public static void Sql(String[] data) throws ClassNotFoundException {
        try {
            Connection connection = SqlConnector.getMySQLConnection();
            Statement statement = connection.createStatement();
            //ResultSet rs = null;

            String sqlInsert = "INSERT INTO `tnk`";
            String columnArray[] = new String[]{"STT", "Mã thiết bị", "Số lượng", "Thời gian bảo hành", "Phần mềm", "column6"};
            String columnNames = "";
            String valueString = "";
            String value[] = data;
            int columnCount = value.length;// get value from excel sheet

            for (int i = 0; i < columnCount; i++) {
                if (i != 0) {
                    columnNames += ",";
                    valueString += ",";
                }
                    columnNames += "`" + columnArray[i] + "`";
                    valueString += value[i];
                
            }
            
            sqlInsert += "(" + columnNames + ") VALUES (" + valueString + ");";

            sqlInsert = sqlInsert.replace("[", "");
            sqlInsert = sqlInsert.replace("]", "");

            //statement.executeUpdate(sqlInsert);
            String checkQuery = "SELECT `Số lượng` AS countValue FROM `tnk` WHERE `Mã thiết bị` = " + value[1] + "";
            statement.execute(checkQuery);
            ResultSet resultSet = statement.getResultSet();
            if (resultSet.next() == true) {
                int Amounts = resultSet.getInt("countValue") + Integer.parseInt(value[2].replace("'", ""));
                String updateQuery = "UPDATE `tnk` SET `Số lượng`=" + String.valueOf(Amounts) + " WHERE `Mã thiết bị`=" + value[1] + ";";
                statement.executeUpdate(updateQuery);
            } else {
                statement.executeUpdate(sqlInsert);
            }
        } catch (SQLException ex) {
            Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public static void main(String args[]) throws IOException, FileNotFoundException, InvalidFormatException, SQLException, ClassNotFoundException {
        FILE_INPUT_STREAM = args[0]; 
        reader();
    }
}
