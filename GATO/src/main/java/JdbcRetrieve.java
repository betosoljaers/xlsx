import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.sql.*;
import java.util.ArrayList;
import java.util.Map;
import java.util.TreeMap;

public class JdbcRetrieve {

    public static void main(String[] args) {

        String url = "jdbc:mysql://localhost:3306/testdb?useUnicode=true&useJDBCCompliantTimezoneShift=true&useLegacyDatetimeCode=false&serverTimezone=UTC";
        String user = "root";
        String password = "";
        ResultSetMetaData resultSetMetaData = null;
        String query = "SELECT title, id FROM Books";

        try {
            Connection con = DriverManager.getConnection(url, user, password);
            PreparedStatement pst = con.prepareStatement(query);
            ResultSet rs = pst.executeQuery();
            resultSetMetaData = rs.getMetaData();

            int numberOfColumns = resultSetMetaData.getColumnCount();
            String[] colum = new String[numberOfColumns];

            TreeMap<String, ArrayList<Object>> listMap = new TreeMap<>();
            for (int i = 1; i < numberOfColumns + 1; i++) {
                String columnName = resultSetMetaData.getColumnName(i);
                colum[i - 1] = columnName;
                listMap.put(columnName, new ArrayList<>());
            }

            while (rs.next()) {
                for (Map.Entry<String, ArrayList<Object>> entry : listMap.entrySet()) {
                    listMap.get(entry.getKey()).add(rs.getString(entry.getKey()));
                }
            }

            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Gato");

            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setFontHeightInPoints((short) 14);
            headerFont.setColor(IndexedColors.RED.getIndex());

            CellStyle headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setFont(headerFont);

            Row headerRow = sheet.createRow(0);

            for (int i = 0; i < colum.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(colum[i]);
                cell.setCellStyle(headerCellStyle);
            }

            Map.Entry<String, ArrayList<Object>> entry = listMap.entrySet().iterator().next();
            int count = entry.getValue().size();
            int temp = 0;
            boolean finish = false;
            Row row;
            Cell cell;

            for (int i = 0; i < colum.length; i++) {
                String c = colum[i];
                ArrayList<Object> list = listMap.get(c);

                for (int j = 0; j < count; j++) {

                    if (finish) {
                        row = sheet.getRow(j + 1);
                    } else {
                        row = sheet.createRow(j + 1);
                    }

                    cell = row.createCell(temp);
                    cell.setCellValue((String) list.get(j));
                }
                temp++;
                finish = true;
            }

            FileOutputStream fileOut = new FileOutputStream("gato.xlsx");
            workbook.write(fileOut);
            fileOut.close();

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
}