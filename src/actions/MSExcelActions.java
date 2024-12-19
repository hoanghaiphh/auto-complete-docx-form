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
    public static final String FILE_NAME = System.getProperty("user.dir") + File.separator
            + "sourceDocs" + File.separator + "THONG_TIN.xlsx";

    public static void addNewRecord(String sheetName, List<String> dataList) {
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(FILE_NAME);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);

            int rowIndex = sheet.getLastRowNum() + 1;
            Row newRow = sheet.createRow(rowIndex);
            newRow.createCell(0).setCellValue(rowIndex);
            for (int i = 1; i < dataList.size(); i++) {
                newRow.createCell(i).setCellValue(dataList.get(i));
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

    public static Row getRowByColumnValue(String sheetName, int columnIndex, String value) {
        Row result = null;
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(FILE_NAME);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row.getCell(columnIndex).getStringCellValue().equalsIgnoreCase(value)) {
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

    public static List<Row> getListRowByColumnValue(String sheetName, int columnIndex, String value) {
        List<Row> result = new ArrayList<>();
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(FILE_NAME);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row.getCell(columnIndex).getStringCellValue().equalsIgnoreCase(value)) {
                    result.add(row);
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

    public static List<String> getListValueOfColumn(String sheetName, int columnIndex) {
        List<String> result = new ArrayList<>();
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(FILE_NAME);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell = row.getCell(columnIndex);
                result.add(cell.getStringCellValue());
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

    public static void updateRecord(String srcFile, String dstFile, String sheetName, int rowIndex, List<String> dataList) {
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(srcFile);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);

            Row row = sheet.getRow(rowIndex);
            for (int i = 1; i < dataList.size(); i++) {
                Cell cell = row.getCell(i);
                if (cell == null) {
                    cell = row.createCell(i);
                }
                cell.setCellValue(dataList.get(i));
            }

            FileOutputStream fos = new FileOutputStream(dstFile);
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
}
