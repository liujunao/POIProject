package poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

/**
 * Created by lenovo on 2017/9/17.
 */
public class FetchExcelToDB {

    //将 Excel 信息存入 Mysql 数据库
    public boolean fetchExcelToMysql(String path, String database) {
        InputStream inputStream = this.getClass().getResourceAsStream("/db.properties");
        Properties properties = new Properties();
        Connection connection = null;
        Statement statement = null;
        try {
            properties.load(inputStream);
            inputStream.close();
            String driver = properties.getProperty("driver").trim();
            String url = properties.getProperty("url");
            String user = properties.getProperty("user");
            String password = properties.getProperty("password");
            Class.forName(driver);
            connection = DriverManager.getConnection(url, user, password);
            statement = connection.createStatement();

            statement.executeUpdate("DROP DATABASE IF EXISTS " + database);
            statement.executeUpdate("CREATE DATABASE  " + database);
            statement.close();
            connection.close();

            List<List<List<String>>> lists = readExcel(path);
            url = "jdbc:mysql://localhost:3306/" + database + "?useSSL=true&serverTimezone=UTC";
            connection = DriverManager.getConnection(url, user, password);
            statement = connection.createStatement();
            for (int i = 0; i < lists.size(); i++) {
                InputStream inputStream1 = new FileInputStream(path);
                //HSSFWorkbook 标识整个 Excel
                Workbook hssfWorkbook = null;
                try {
                    hssfWorkbook = new HSSFWorkbook(inputStream1);
                } catch (Exception e) {
                    e.printStackTrace();
                    inputStream1 = new FileInputStream(path);
                    hssfWorkbook = new XSSFWorkbook(inputStream1);
                }
                String dataTable = hssfWorkbook.getSheetName(i);
                if (hasDataTable(dataTable,database)) {
                    statement.executeUpdate("DROP TABLE " + dataTable);
                }
                String dataTableId = dataTable + "Id";
                statement.executeUpdate("CREATE TABLE " + dataTable + "(" + dataTableId + " INT NOT NULL AUTO_INCREMENT," +
                        "PRIMARY KEY (" + dataTableId + ")" + ")ENGINE=InnoDB DEFAULT CHARSET=utf8");

                List<List<String>> lists1 = lists.get(i);
                List<String> lists2 = lists1.get(0);
                //添加相应的字段
                for (int j = 1; j < lists2.size(); j++) {
                    String name = lists2.get(j);
                    statement.executeUpdate("ALTER TABLE " + dataTable + " ADD " + name + " VARCHAR(225)");
                }
                String insertName = "";
                for (int m = 1; m < lists2.size() - 1; m++) {
                    insertName += lists2.get(m) + ",";
                }
                insertName += lists2.get(lists2.size() - 1);
                for (int k = 1; k < lists1.size(); k++) {
                    List<String> result = lists1.get(k);
                    String sql = "INSERT INTO " + dataTable + "(" + insertName + ")VALUES (";
                    for (int t = 0; t < result.size() - 1; t++) {
                        sql += "'" + result.get(t) + "'" + ",";
                    }
                    String lastList = result.get(result.size() - 1);
                    sql += "'" + lastList + "'" + ")";
                    statement.executeUpdate(sql);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
           return false;
        }
        return true;
    }

    //读取 Excel 内容
    public List<List<List<String>>> readExcel(String path) {
        List<List<List<String>>> result = new ArrayList<>();
        try {
            InputStream inputStream = new FileInputStream(path);
            //HSSFWorkbook 标识整个 Excel
            Workbook hssfWorkbook = null;
            try {
                hssfWorkbook = new HSSFWorkbook(inputStream);
            } catch (Exception e) {
                inputStream = new FileInputStream(path);
                hssfWorkbook = new XSSFWorkbook(inputStream);
            }
            int size = hssfWorkbook.getNumberOfSheets();

            for (int numberSheet = 0; numberSheet < size; numberSheet++) {
                //HSSFSheet 标识某一页
                Sheet hssfSheet = hssfWorkbook.getSheetAt(numberSheet);
                List<List<String>> sheetList = new ArrayList<>();
                if (hssfSheet == null) {
                    continue;
                }
                for (int rowNumber = 0; rowNumber <= hssfSheet.getLastRowNum(); rowNumber++) {
                    Row hssfRow = hssfSheet.getRow(rowNumber);
                    int minRowCellIndex = hssfRow.getFirstCellNum();
                    int maxRowCellIndex = hssfRow.getLastCellNum();
                    List<String> rowList = new ArrayList<>();
                    for (int cellIndex = minRowCellIndex; cellIndex < maxRowCellIndex; cellIndex++) {
                        Cell hssfCell = hssfRow.getCell(cellIndex);
                        if (hssfCell == null) {
                            continue;
                        }
                        rowList.add(getStringValue(hssfCell));
                    }
                    sheetList.add(rowList);
                }
                result.add(sheetList);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return result;
    }

    //进行 cell 的数据类型转换
    private String getStringValue(Cell hssfCell) {
        switch (hssfCell.getCellType()) {
            case Cell.CELL_TYPE_BOOLEAN:
                return hssfCell.getBooleanCellValue() ? "TRUE" : "FALSE";
            case Cell.CELL_TYPE_FORMULA:
                return hssfCell.getCellFormula();
            case Cell.CELL_TYPE_NUMERIC:
                hssfCell.setCellType(Cell.CELL_TYPE_STRING);
                return hssfCell.getStringCellValue();
            case Cell.CELL_TYPE_STRING:
                return hssfCell.getStringCellValue();
            default:
                return "";
        }
    }

    //判断是否存在相应的数据表
    public boolean hasDataTable(String dataTable,String database) {

        boolean flag = false;
        Common common = new Common();
        Connection connection = common.getConnection();
        try {
            ResultSet resultSet = connection.getMetaData().getTables(database, null, dataTable, null);
            if (resultSet.next()) {
                flag = true;
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return flag;
    }

}