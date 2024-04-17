import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelToText {
    public static void main(String[] args) {
        try {
            //创建读取路径：
            String PATH = "C:\\Users\\Yzy\\IdeaProjects\\dormitory\\Data\\";
            //获取文件流：
            FileInputStream fis = new FileInputStream(PATH + "屏峰校区查寝分数统计表第十七周.xlsx");
            Workbook workbook = new XSSFWorkbook(fis);
            DataFormat format = workbook.createDataFormat();

            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setDataFormat(format.getFormat("@"));

            for (Sheet sheet : workbook) {
                for (Row row : sheet) {
                    if (row == null) break;
                    for (Cell cell : row) {
                        if (cell == null) break;
                        cell.setCellStyle(cellStyle);
                    }
                }
            }

            fis.close();

            FileOutputStream fos = new FileOutputStream(PATH + "屏峰校区查寝分数统计表第十七周.xlsx");
            workbook.write(fos);
            fos.close();
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}