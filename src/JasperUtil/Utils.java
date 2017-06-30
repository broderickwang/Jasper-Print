package JasperUtil;

import java.sql.Connection;
import java.sql.DriverManager;

/*
 *Created by Broderick
 * User: Broderick
 * Date: 2017/6/23
 * Time: 15:02
 * Version: 1.0
 * Description:
 * Email:wangchengda1990@gmail.com
**/
public class Utils {

    public static Connection getConnection() throws Exception{
        Class.forName("oracle.jdbc.driver.OracleDriver");
        String url = "jdbc:oracle:" + "thin:@192.168.9.202:1521:orcl";
        String user = "sde";
        String password = "sde";
        Connection conn = DriverManager.getConnection(url, user, password);

        return conn;
    }
    public static Connection getConnection(String ip,String orcl,String user,String password) throws Exception{
        Class.forName("oracle.jdbc.driver.OracleDriver");
        String url = "jdbc:oracle:" + "thin:@"+ip+":"+orcl;
        Connection conn = DriverManager.getConnection(url, user, password);

        return conn;
    }
}
