
import entity.Normal_Student;
import mapper.InsertMapper;
import org.apache.ibatis.session.SqlSession;
import org.apache.ibatis.session.SqlSessionFactory;
import org.apache.ibatis.session.SqlSessionFactoryBuilder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;


public class dormitoryScore {
    public static void main(String[] args) throws IOException {
        SqlSessionFactory build = new SqlSessionFactoryBuilder().build(new FileInputStream("src/main/resources/mybatis-config.xml"));
        try (SqlSession session = build.openSession(true)) {

            //创建读取路径：
            String PATH = "C:\\Users\\Yzy\\IdeaProjects\\dormitory\\Data\\";
            //获取文件流：
            FileInputStream inputStream = new FileInputStream(PATH + "屏峰校区查寝分数统计表第十七周.xlsx");
            Workbook workbook = new XSSFWorkbook(inputStream);
            // 获取工作簿中的所有 sheet
            int numberOfSheets = workbook.getNumberOfSheets();

            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();
                // 从第三行开始遍历当前 sheet 的每一行
                System.out.println("开始读取表："+sheetName);
                for (int j = 2; j <= sheet.getLastRowNum(); j++) {
                    Row row = sheet.getRow(j);
                    if (row != null) {
                        // 读取当前行的四个单元格内容并保存到变量中
                        Cell cell1 = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        Cell cell2 = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        Cell cell3 = row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        Cell cell4 = row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        ArrayList<Cell> cells=new ArrayList<>();
                        cells.add(cell1);
                        cells.add(cell2);
                        cells.add(cell3);
                        cells.add(cell4);
                        ArrayList<String> strings=new ArrayList<>();
                        for(Cell cell :cells){
                            if (cell != null) {
                                switch (cell.getCellType()) {
                                    case STRING:
                                        strings.add(cell.getStringCellValue());
                                        break;
                                    case NUMERIC:
                                        strings.add(String.valueOf(cell.getNumericCellValue()));
                                        break;
                                    default:
                                        System.out.println("Invalid cell type");
                                }
                            }
                        }

                        if(strings.size()<4){
                            continue;
                        }

                        if(strings.get(3).equals("计算机科学与技术学院、软件学院")) {
                            InsertMapper mapper = session.getMapper(InsertMapper.class);  //通过接口获取实现类
                            Normal_Student student=new Normal_Student();
                            student.setBuilding(sheetName);
                            student.setRoom_number(strings.get(0));
                            student.setBed_number(strings.get(1));
                            student.setScore(strings.get(2));
                            System.out.println(mapper.addNormalStudent(student));
                        }
                    }
                }
            }
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
