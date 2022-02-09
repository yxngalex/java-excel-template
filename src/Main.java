import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    public static void main(String[] args) {
        String excelFilePath = "eli.xlsx";

        try {
            FileInputStream inputStream = new FileInputStream(excelFilePath);
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheetAt(0);

            for (int i = 3; i < sheet.getPhysicalNumberOfRows(); i++) {
                XSSFRow row = sheet.getRow(i);
                String cdl;
                if (row.getCell(1).getCellType() == CellType.STRING) {
                    cdl = row.getCell(1).getStringCellValue();
                } else {
                    cdl = String.valueOf(row.getCell(1).getNumericCellValue());
                }
                String issueProvince = row.getCell(2).getStringCellValue();

                if (issueProvince.equals("ON")) {
                    cdl = fixCdl(cdl);
                    XSSFCell cell = row.getCell(1);
                    cell.setCellValue(cdl);
                }
            }

            inputStream.close();

            FileOutputStream outputStream = new FileOutputStream(excelFilePath);

            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String fixCdl(String cdl) {
        StringBuilder builder = new StringBuilder();
        char[] arr = cdl.toCharArray();

        for (int i = 0; i < arr.length; i++) {
            builder.append(arr[i]);
            if ((i + 1) % 5 == 0 && i != arr.length - 1) {
                builder.append("-");
            }
        }

        return builder.toString();
    }
}
