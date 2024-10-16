package actions;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class MSExcelActions {
    private static final String FILE_NAME = System.getProperty("user.dir") + File.separator
            + "sourceDocs" + File.separator + "THONG_TIN.xlsx";

    public static void addNewRecord(String sheetName, List<String> authorizerInfo) {
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(FILE_NAME);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);

            int rowIndex = sheet.getLastRowNum() + 1;
            Row newRow = sheet.createRow(rowIndex);
            newRow.createCell(0).setCellValue(rowIndex);
            for (int i = 1; i < authorizerInfo.size(); i++) {
                newRow.createCell(i).setCellValue(authorizerInfo.get(i));
            }

            FileOutputStream fos = new FileOutputStream(FILE_NAME);
            workbook.write(fos);
            fos.close();
            workbook.close();

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (fis != null) {
                    fis.close();
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    public static List<String> getAllAuthorizerComIds() {
        List<String> comIds = new ArrayList<>();
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(FILE_NAME);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet("UY_QUYEN");
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell = row.getCell(3);
                comIds.add(cell.getStringCellValue());
            }
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (fis != null) {
                    fis.close();
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return comIds;
    }

    public static Row getRowByAuthorizerComId(String authorizerComId) {
        Row result = null;
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(FILE_NAME);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet("UY_QUYEN");
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row.getCell(3).getStringCellValue().equals(authorizerComId)) {
                    result = row;
                    break;
                }
            }
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (fis != null) {
                    fis.close();
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return result;
    }
}
