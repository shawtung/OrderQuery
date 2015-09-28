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
import jxl.write.Number;
import jxl.write.biff.RowsExceededException;


public class XLS {
    // JDBC driver name and database URL
    static final String DB_URL_TEST = "jdbc:mysql://121.41.33.100:3306/51qed?useUnicode=true&characterEncoding=UTF-8";
    static final String DB_URL_PROD = "jdbc:mysql://121.41.106.210:3306/51qed?useUnicode=true&characterEncoding=UTF-8";

    //  Database credentials
    static final String USER_TEST = "root";
    static final String PASS_TEST = "";
    static final String USER_PROD = "samewayread";
    static final String PASS_PROD = "sameway25@*&1";

	static String SQL;
	static String DB_URL, USER, PASS;
	static File file;

	public static void Switch(String env, String sql, String date) {
		switch (env) {
			case "PROD" : {
				DB_URL = DB_URL_PROD;
				USER = USER_PROD;
				PASS = PASS_PROD;
				break;
			}

			case "TEST" : {
				DB_URL = DB_URL_TEST;
				USER = USER_TEST;
				PASS = PASS_TEST;
				break;
			}
		}

		switch (sql) {
			case "OrderQuerySQL" : {
				SQL = "SELECT\n" +
						"    order_info.id AS 订单ID,\n" +
						"    product_name AS 项目名称,\n" +
						"    user_info.real_name AS 用户姓名,\n" +
						"    user_info.telephone AS 手机号,\n" +
						"    ATT.province AS 归属地,\n" +
						"    register_channel AS 注册渠道,\n" +
						"    identity_card AS 身份证号,\n" +
						"    IF(SUBSTR(identity_card, -2, 1) IN (1, 3, 5, 7, 9), '男', '女') AS 性别,\n" +
						"    ROUND(order_info.order_money) AS 订单金额（元）,\n" +
						"    DATE_FORMAT(order_info.create_time, '%Y-%c-%e %T') AS 订单创建时间,\n" +
						"    DATE_FORMAT(order_info.pay_time, '%Y-%c-%e %T') AS 订单确认时间,\n" +
						"    COUNT(order_info.id) AS 个人投资序数,\n" +
						"    ROUND(`SUM(order_money)`) AS 个人投资总额（元）,\n" +
						"    IF(COUNT(order_info.id) = 1, '是', '') AS 是否首投,\n" +
						"    IFNULL(H.`real_name`, '') AS 推荐人,\n" +
						"    IFNULL(HH.`real_name`, '') AS 二级推荐人\n" +
						"FROM\n" +
						"    order_info\n" +
						"        INNER JOIN\n" +
						"    product_info ON order_info.order_product_tid = product_info.id AND order_info.order_status_id = 100567\n" +
						"        INNER JOIN\n" +
						"    order_info AS OI2 ON OI2.order_user_tid = order_info.order_user_tid AND order_info.id >= OI2.id AND OI2.order_status_id = 100567\n" +
						"        INNER JOIN\n" +
						"    user_info ON order_info.order_user_tid = user_info.id\n" +
						"        LEFT JOIN\n" +
						"    (SELECT telephone, province FROM user_attribute_info WHERE province IN ('上海' , '北京', '天津', '重庆', '安徽', '浙江', '江苏', '湖北', '湖南', '河北', '河南', '福建', '广东', '广西', '台湾', '贵州', '云南', '四川', '西藏', '新疆', '青海', '陕西', '山西', '山东', '宁夏', '甘肃', '辽宁', '吉林', '江西', '海南', '内蒙古', '黑龙江', '香港', '澳门') GROUP BY telephone) AS ATT ON ATT.telephone = user_info.telephone\n" +
						"        LEFT JOIN\n" +
						"    (SELECT user_info.id, user_info.real_name, user_info.referrer_id FROM user_info) AS H ON user_info.referrer_id = H.id\n" +
						"        LEFT JOIN\n" +
						"    (SELECT user_info.id, user_info.real_name, user_info.referrer_id FROM user_info) AS HH ON H.referrer_id = HH.id\n" +
						"        LEFT JOIN\n" +
						"    (SELECT order_user_tid, SUM(order_money) FROM order_info WHERE order_status_id = 100567 GROUP BY order_user_tid) AS X ON X.order_user_tid = order_info.order_user_tid\n" +
						"WHERE order_info.create_time >= '" + date + "'\n" +
						"GROUP BY order_info.id\n" +
						"ORDER BY order_info.create_time DESC;";
				file = new File("./" + date + "至" + (new SimpleDateFormat("yyyy-MM-dd")).format(new Date()) + "投资订单报表.xls");
				break;
			}

			case "WithdrawQuerySQL" : {
                SQL = "SELECT\n" +
		                "    user_info.real_name AS 用户姓名,\n" +
		                "    user_info.telephone AS 用户手机号,\n" +
		                "    DATE_FORMAT(withDraw_application.create_time, '%Y-%c-%e %T') AS 提现申请时间,\n" +
		                "    ROUND(withDraw_application.amount, 2) AS 提现金额,\n" +
		                "    withDraw_application.counter_fee AS 手续费金额,\n" +
		                "    ROUND((withDraw_application.amount - withDraw_application.counter_fee), 2) AS 应结算金额,\n" +
		                "    bank_info.bank_name AS 银行名称,\n" +
		                "    bank_info.subbranch_name AS 支行名称,\n" +
		                "    bank_info.account_info AS 银行账号,\n" +
		                "    withDraw_application.audit_date AS 审核日期\n" +
		                "FROM\n" +
		                "    withDraw_application,\n" +
		                "    user_info,\n" +
		                "    bank_info\n" +
		                "WHERE\n" +
		                "    withDraw_application.create_date >= '" + date + "'\n" +
		                "        AND withDraw_application.user_tid = user_info.id\n" +
		                "        AND user_info.user_account_tid = bank_info.id\n" +
		                "ORDER BY withDraw_application.create_time DESC;";
				file = new File("./" + date + "至" + (new SimpleDateFormat("yyyy-MM-dd")).format(new Date()) + "提现报表.xls");
				break;
			}

			case "AuthenticatedUserQuerySQL" : {
                SQL = "SELECT\n" +
		                "    user_info.real_name AS 用户姓名,\n" +
		                "    user_info.telephone AS 手机号,\n" +
		                "    ATT.province AS 归属地,\n" +
		                "    CONCAT('\u200B', identity_card) AS 身份证号,\n" +
		                "    IFNULL(T.`real_name`, '') AS 推荐人,\n" +
		                "    IFNULL(TT.`real_name`, '') AS 二级推荐人,\n" +
		                "    DATE_FORMAT(user_info.create_time, '%Y-%c-%e %T') AS 注册时间\n" +
		                "FROM\n" +
		                "    user_info\n" +
		                "        LEFT JOIN\n" +
		                "    (SELECT telephone, province FROM user_attribute_info WHERE province IN ('上海' , '北京', '天津', '重庆', '安徽', '浙江', '江苏', '湖北', '湖南', '河北', '河南', '福建', '广东', '广西', '台湾', '贵州', '云南', '四川', '西藏', '新疆', '青海', '陕西', '山西', '山东', '宁夏', '甘肃', '辽宁', '吉林', '江西', '海南', '内蒙古', '黑龙江', '香港', '澳门') GROUP BY telephone) AS ATT ON ATT.telephone = user_info.telephone\n" +
		                "        LEFT JOIN\n" +
		                "    (SELECT id, real_name, referrer_id FROM user_info) AS T ON T.id = user_info.referrer_id\n" +
		                "        LEFT JOIN\n" +
		                "    (SELECT id, real_name, referrer_id FROM user_info) AS TT ON T.referrer_id = TT.id\n" +
		                "WHERE\n" +
		                "    create_time >= '" + date + "'\n" +
		                "        AND identity_card IS NOT NULL\n" +
		                "ORDER BY user_info.create_time DESC;";
				file = new File("./" + date + "至" + (new SimpleDateFormat("yyyy-MM-dd")).format(new Date()) + "认证用户报表.xls");
				break;
			}
		}
	}

    public static void makeXLS(String env, String sql, String date) {
        Connection conn = null;
        Statement stmt = null;

	    Switch(env, sql, date);

        try {
            //Register JDBC driver
            Class.forName("com.mysql.jdbc.Driver");

            //Open a connection
            System.out.println("Connecting to database...");
            conn = DriverManager.getConnection(DB_URL, USER, PASS);

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
            WritableCellFormat headerFormat = new WritableCellFormat(new WritableFont(WritableFont.createFont("宋体"), 10, WritableFont.BOLD));
            do {
                for (int j = 1; j <= rsmd.getColumnCount(); j++) {
                    if (i == 0) {
                        sheet.addCell(new Label(j - 1, i, rsmd.getColumnLabel(j), headerFormat));
                    } else {
	                    if (rsmd.getColumnLabel(j).indexOf("金额") > -1) {
		                    sheet.addCell(new Number(j - 1, i, Double.parseDouble(rs.getString(j)), new WritableCellFormat(NumberFormats.FLOAT)));
	                    } else {
		                    sheet.addCell(new Label(j - 1, i, rs.getString(j)));
	                    }
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
            }
        }
	    System.out.println("Finished\n");
    }

	public static void main(String[] args) {
		makeXLS("PROD", "WithdrawQuerySQL", "2015-09-26");

	}
}
