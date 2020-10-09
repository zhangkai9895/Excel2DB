import java.sql.*;
import java.util.ArrayList;
import java.util.Map;

public class Excel2DB {
        private static final String URL="jdbc:mysql://localhost:3306/student?useSSL=false&serverTimezone=UTC&allowPublicKeyRetrieval=true";
        private static final String NAME="root";
        private static final String PASSWORD="0122";
        private static PreparedStatement ps=null;

        public static void main(String[] args) throws Exception{

            //1.加载驱动程序
            Class.forName("com.mysql.jdbc.Driver");
            //2.获得数据库的连接
            Connection conn = DriverManager.getConnection(URL, NAME, PASSWORD);
            //3.通过数据库的连接操作数据库，实现增删改查

            String sql="insert into course (id,coursecode,coursename,coursetype,coursehours,coursecredit,peoplenum,coursecollege,courseteacher,teachertype,weekperiod,jointime,courseperiod,location,weekday) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
            ps=conn.prepareStatement(sql);

            ExcelReader excelReader =new ExcelReader();
            ArrayList<Map<String,String>> result = excelReader.readExcelToObj("C:\\Users\\zhangkai\\Desktop\\test.xls");
            for(Map<String,String> map:result){
                ps.setLong(1, Integer.parseInt(map.get("id")));
                ps.setString(2, map.get("code"));
                ps.setString(3, map.get("courseName"));
                ps.setString(4,"沟通与管理（原 人文社会科学类）" );
                ps.setInt(5,Integer.parseInt(map.get("courseHours")));
                ps.setInt(6,Integer.parseInt(map.get("courseCredit")));
                ps.setInt(7,Integer.parseInt(map.get("peopleNum")));
                ps.setString(8,map.get("courseCollege"));
                ps.setString(9,map.get("courseTeacher"));
                ps.setString(10,map.get("teacherType"));
                ps.setString(11,map.get("weekPeriod"));
                ps.setString(12,map.get("coursePeriod"));
                ps.setString(13,map.get("location"));
                ps.setString(14,map.get("weekday"));
                ps.setString(15,null);
                ps.executeUpdate();
            }


        }

}
