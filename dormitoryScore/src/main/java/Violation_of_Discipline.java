import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;

public class Violation_of_Discipline {
    public static void main(String[] args) {
        try {

            //创建读取路径：
            String PATH = "C:\\Users\\Yzy\\IdeaProjects\\dormitory\\Data\\";
            //获取文件流：
            FileInputStream inputStream = new FileInputStream(PATH + "第十二周.xlsx");
            Workbook workbook = new XSSFWorkbook(inputStream);
            String jdbcUrl = "jdbc:mysql://localhost:3306/Violation";
            String user = "root";
            String password = "root";

            Connection conn = DriverManager.getConnection(jdbcUrl, user, password);
            // SQL 插入语句，假设表名为 example_table，字段分别为 column1, column2, column3

            String sql = "INSERT INTO twelve (name,room_number, time, reason, student_number) VALUES (?, ?, ?, ?, ?)";
            // 创建 PreparedStatement 对象
            PreparedStatement preparedStatement = conn.prepareStatement(sql);

            Sheet sheet = workbook.getSheetAt(0);
            // 从第三行开始遍历当前 sheet 的每一行
            for (int j = 1; j <= sheet.getLastRowNum(); j++) {
                Row row = sheet.getRow(j);
                if (row != null) {
                    // 读取当前行的四个单元格内容并保存到变量中
                    Cell cell1 = row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    Cell cell2 = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    Cell cell3 = row.getCell(13, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    Cell cell4 = row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    Cell cell5 = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);


                    String value1 = (String) getValueFromCell(cell1);
                    String value2 = (String) getValueFromCell(cell2);
                    String value3 = (String) getValueFromCell(cell3);
                    String value4 = (String) getValueFromCell(cell4);
                    String value5 = (String) getValueFromCell(cell5);

                    // 设置参数
                    preparedStatement.setString(1, value1); // 设置第一个参数的值
                    preparedStatement.setString(2, value2); // 设置第二个参数的值
                    preparedStatement.setString(3, value3); // 设置第三个参数的值
                    preparedStatement.setString(4, value4); // 设置第四个参数的值
                    preparedStatement.setString(5, value5); // 设置第五个参数的值

                    // 执行 SQL 语句
                    preparedStatement.executeUpdate();
                }
            }
            inputStream.close();
            preparedStatement.close();
        } catch (SQLException ex) {
            throw new RuntimeException(ex);
        } catch (FileNotFoundException ex) {
            throw new RuntimeException(ex);
        } catch (IOException ex) {
            throw new RuntimeException(ex);
        }

    }

    private static Object getValueFromCell(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                double numericValue = cell.getNumericCellValue();
                return numericValue;
            default:
                return ""; // 对于其他类型的单元格，返回空字符串
        }
    }
}
