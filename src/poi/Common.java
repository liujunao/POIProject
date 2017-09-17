package poi;

import java.io.*;
import java.sql.*;
import java.util.*;

/**
 * Created by lenovo on 2017/9/17.
 */
//修改 properties 属性文件
public class Common {

    public void revise(String path,String key,String value){
        Properties properties = new Properties();
        InputStream inputStream = null;
        OutputStream outputStream = null;
        File file = new File(path);
        try {
            inputStream = new FileInputStream(file);
            properties.load(inputStream);
            inputStream.close();
            outputStream = new FileOutputStream(file);
            //设值，保存
            properties.setProperty(key,value);
            properties.store(outputStream,null);
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    //获取连接
    public Connection getConnection() {
        Properties properties = new Properties();
        String dirPath = Thread.currentThread().getContextClassLoader().getResource("").getPath();
        System.out.println(dirPath);
        Connection connection = null;
        Statement statement = null;
        try {
            InputStream inputStream = new FileInputStream(dirPath + "/db.properties");
            properties.load(inputStream);
            inputStream.close();
            String driver = properties.getProperty("driver").trim();
            String url = properties.getProperty("url");
            String user = properties.getProperty("user");
            String password = properties.getProperty("password");
            Class.forName(driver);
            connection = DriverManager.getConnection(url, user, password);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return connection;
    }

    //查询数据
    public List<Map<String,Object>> query(String dataTable,Object ... args){
        List<Map<String,Object>> list = null;
        ResultSet resultSet = null;
        Connection connection = null;
        PreparedStatement preparedStatement = null;
        connection = getConnection();
        String sql = "SELECT * FROM " + dataTable;
        try {
            preparedStatement = connection.prepareStatement(sql);
            if (args != null && args.length > 0){
                for (int i = 0;i < args.length;i++){
                    preparedStatement.setObject(i + 1,args[i]);
                }
            }
            resultSet = preparedStatement.executeQuery();

            if (resultSet != null){
                list = getResultMap(resultSet);
                connection.close();
                preparedStatement.close();
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }

        return list;
    }

    //将 Result 转换为 List<Map<String,Object>> 类型数据
    private List<Map<String,Object>> getResultMap(ResultSet resultSet) {
        Map<String,Object> map = null;
        List<Map<String,Object>>list = new ArrayList<>();
        try {
            ResultSetMetaData resultSetMetaData = resultSet.getMetaData();
            int count = resultSetMetaData.getColumnCount();
            while (resultSet.next()){
                map = new HashMap<>();
                for (int i = 1;i <= count;i++){
                    String key = resultSetMetaData.getColumnLabel(i);
                    Object value = resultSet.getString(i);
                    map.put(key,value);
                }
                list.add(map);
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return list;
    }

    //获取数据表的列名
    public List<String> queryColumnName(String dataTable){
        List<String> columnName = new ArrayList<>();
        Connection connection = null;
        PreparedStatement preparedStatement = null;
        ResultSet resultSet = null;
        connection = getConnection();

        String sql = "SELECT * FROM " + dataTable;
        try {
            preparedStatement = connection.prepareStatement(sql);
            resultSet = preparedStatement.executeQuery();
            ResultSetMetaData resultSetMetaData = resultSet.getMetaData();
            int count = resultSetMetaData.getColumnCount();
            for (int i = 1;i <= count;i++){
                columnName.add(resultSetMetaData.getColumnName(i));
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return columnName;
    }

    //判断文件是否存在
    public boolean fileExist(String fileDir){
        boolean flag = false;
        File file = new File(fileDir);
        flag = file.exists();

        return flag;
    }

    //删除文件
    public boolean deleteExcel(String fileDir){
        boolean flag = false;
        File file = new File(fileDir);
        //判断目录或文件是否存在
        if (!file.exists()){
            return false;
        }else {
            //判断是否为文件
            if (file.isFile()){
                file.delete();
                flag = true;
            }
        }
        return flag;
    }

}
