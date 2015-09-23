/**
 * Created by Shaw on 2015/9/3.
 */

//STEP 1. Import required packages

import java.sql.*;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;

import jxl.*;
import jxl.write.*;
import jxl.write.biff.RowsExceededException;


public class XLS {
    // JDBC driver name and database URL
    static final String DBT_URL = "jdbc:mysql://121.41.33.100:3306/51qed?useUnicode=true&characterEncoding=UTF-8";
    static final String DB_URL = "jdbc:mysql://121.41.106.210:3306/51qed?useUnicode=true&characterEncoding=UTF-8";

    //  Database credentials
    static final String USERT = "root";
    static final String PASST = "";
    static final String USER = "samewayread";
    static final String PASS = "sameway25@*&1";
	static String SQL;
	static String OrderQuerySQL;
	static String WithdrawQuerySQL;
	static String AuthenticatedUserQuerySQL;
	static File file;

	public static void Switch(String sql, String date) {
		switch (sql) {
			case "OrderQuerySQL" : {
				SQL = "SELECT \n" +
						"    product_info.product_name AS '项目名称',\n" +
						"    user_info.real_name AS '用户姓名',\n" +
						"    user_info.telephone AS '手机号',\n" +
						"    ROUND(order_info.order_money) AS '订单金额',\n" +
						"    Date_FORMAT(order_info.create_time, '%Y-%m-%d %T') AS '订单创建时间',\n" +
						"    Date_FORMAT(order_info.pay_time, '%Y-%m-%d %T') AS '订单确认时间',\n" +
						"    Z.ord AS '个人序数',\n" +
						"    IF(Z.ord = 1, '\u662f', '\u5426') AS '是否首投',\n" +
						"    ifnull(H.real_name, '') AS '推荐人'\n" +
						"FROM\n" +
						"    (SELECT @ord:=0, @pre_uid:=- 1) AS r,\n" +
						"    order_info\n" +
						"        INNER JOIN\n" +
						"    product_info ON order_info.order_product_tid = product_info.id\n" +
						"        AND order_info.order_status_id = 100567\n" +
						"        INNER JOIN\n" +
						"    user_info ON order_info.order_user_tid = user_info.id\n" +
						"        LEFT JOIN\n" +
						"    (SELECT \n" +
						"        user_info.id AS id, user_info.real_name AS real_name\n" +
						"    FROM\n" +
						"        user_info) AS H ON user_info.referrer_id = H.id\n" +
						"        LEFT JOIN\n" +
						"    (SELECT \n" +
						"        id,\n" +
						"            CASE\n" +
						"                WHEN @pre_uid = order_user_tid THEN @ord:=@ord + 1\n" +
						"                WHEN @pre_uid != order_user_tid THEN @ord:=1\n" +
						"            END AS ord,\n" +
						"            @pre_uid:=order_user_tid\n" +
						"    FROM\n" +
						"        order_info, (SELECT @ord, @pre_uid) AS r\n" +
						"    WHERE\n" +
						"        order_status_id = 100567\n" +
						"    ORDER BY order_user_tid , create_time) AS Z ON Z.id = order_info.id\n" +
						"WHERE\n" +
						"    order_info.create_time >= '" + date + "'\n" +
						"ORDER BY order_info.create_time DESC;";
				file = new File("./" + date + "至" + (new SimpleDateFormat("yyyy-MM-dd")).format(new Date()) + "投资订单报表.xls");
			}

			case "WithdrawQuerySQL" : {

				file = new File("./" + date + "至" + (new SimpleDateFormat("yyyy-MM-dd")).format(new Date()) + "提现报表.xls");
			}

			case "AuthenticatedUserQuerySQL" : {

				file = new File("./" + date + "至" + (new SimpleDateFormat("yyyy-MM-dd")).format(new Date()) + "认证用户报表.xls");
			}

			default : {

			}
		}

	}



    public static void makeXLS(String sql, String date) {
        Connection conn = null;
        Statement stmt = null;

	    Switch(sql, date);

        try {
            //Register JDBC driver
            Class.forName("com.mysql.jdbc.Driver");

            //Open a connection
            System.out.println("Connecting to database...");
            conn = DriverManager.getConnection(DBT_URL, USERT, PASST);

            //Execute a query
            System.out.println("Creating statement...");
            stmt = conn.createStatement();
            ResultSet rs = stmt.executeQuery(SQL);
            ResultSetMetaData rsmd = rs.getMetaData();

            //Extract data from result set into Excel
            System.out.println("Creating .xls file...");

            WritableWorkbook wwb = Workbook.createWorkbook(file);
            WritableSheet sheet = wwb.createSheet(date, 0);

            int i = 0;
            WritableFont wFont = new WritableFont(WritableFont.createFont("宋体"), 10, WritableFont.BOLD);
            WritableCellFormat headerFormat = new WritableCellFormat(wFont);
            do {
                for (int j = 1; j <= rsmd.getColumnCount(); j++) {
                    if (i == 0) {
                        sheet.addCell(new Label(j - 1, i, rsmd.getColumnLabel(j), headerFormat));
                    } else {
                        sheet.addCell(new Label(j - 1, i, rs.getString(j)));
                    }
                }
            } while (rs.next() && (++i != 0));

            //Write in workbook and clean-up environment
            wwb.write();
            wwb.close();
            rs.close();
        } catch (SQLException se) {
            //Handle errors for JDBC
            se.printStackTrace();
        } catch (IOException ie) {
            //Handle errors for IO
            ie.printStackTrace();
        } catch (RowsExceededException ree) {
            //Handle errors for sheet
            ree.printStackTrace();
        } catch (WriteException we) {
            //Handle errors for write&close
            we.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            //finally block used to close resources
            try {
                if (stmt != null) {
                    stmt.close();
                    System.out.println("Statement closed...");
                }
            } catch (SQLException se2) {
                se2.printStackTrace();
            }
            try {
                if (conn != null) {
                    conn.close();
                    System.out.println("Connection closed...");
                }
            } catch (SQLException se) {
                se.printStackTrace();
            }//end finally
            System.out.println("Finished");
        }//end try
    }//end make
}
