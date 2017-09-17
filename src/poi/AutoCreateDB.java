package poi;

import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Properties;

/**
 * Created by lenovo on 2017/9/17.
 */
public class AutoCreateDB {

    public void createDB(){
        InputStream inputStream = this.getClass().getClassLoader().getResourceAsStream("/db.properties");
        Properties properties = new Properties();
        Connection connection = null;
        Statement statement = null;
        try {
            properties.load(inputStream);
            String driver = properties.getProperty("driver").trim();
            String url = properties.getProperty("url");
            String user = properties.getProperty("user");
            String password = properties.getProperty("password");
            Class.forName(driver);
            connection = DriverManager.getConnection(url,user,password);
            statement = connection.createStatement();


        } catch (IOException e) {
            e.printStackTrace();
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }
}
