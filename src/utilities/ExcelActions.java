package utilities;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import static utilities.StringConverter.removeDiacritics;

public class ExcelActions {
    public static final String FILE_NAME = System.getProperty("user.dir") + File.separator
            + "sourceDocs" + File.separator + "THONG_TIN.xlsx";

    public static void addNewRecord(String sheetName, List<String> dataList) {
        try (FileInputStream fis = new FileInputStream(FILE_NAME);
             Workbook workbook = new XSSFWorkbook(fis);
             FileOutputStream fos = new FileOutputStream(FILE_NAME)) {

            Sheet sheet = workbook.getSheet(sheetName);
            int rowIndex = sheet.getLastRowNum() + 1;
            Row newRow = sheet.createRow(rowIndex);
            newRow.createCell(0).setCellValue(rowIndex);
            for (int i = 1; i < dataList.size(); i++) {
                newRow.createCell(i).setCellValue(dataList.get(i));
            }

            workbook.write(fos);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static Row getRowByColumnValue(String sheetName, int columnIndex, String value) {
        Row result = null;
        try (FileInputStream fis = new FileInputStream(FILE_NAME);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(sheetName);
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell = row.getCell(columnIndex);
                if (cell != null && cell.getStringCellValue().equalsIgnoreCase(value)) {
                    result = row;
                    break;
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    public static List<Row> getListRowByColumnValue(String sheetName, int columnIndex, String value) {
        List<Row> result = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(FILE_NAME);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(sheetName);
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell = row.getCell(columnIndex);
                if (cell != null && cell.getStringCellValue().equalsIgnoreCase(value)) {
                    result.add(row);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    public static List<String> getListValueOfColumn(String sheetName, int columnIndex) {
        List<String> result = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(FILE_NAME);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(sheetName);
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell = row.getCell(columnIndex);
                if (cell != null) {
                    result.add(cell.getStringCellValue());
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    public static void updateRecord(String sheetName, int rowIndex, List<String> dataList) {
        try (FileInputStream fis = new FileInputStream(FILE_NAME);
             Workbook workbook = new XSSFWorkbook(fis);
             FileOutputStream fos = new FileOutputStream(FILE_NAME)) {

            Sheet sheet = workbook.getSheet(sheetName);
            Row row = sheet.getRow(rowIndex);
            for (int i = 1; i < dataList.size(); i++) {
                Cell cell = row.getCell(i);
                if (cell == null) {
                    cell = row.createCell(i);
                }
                cell.setCellValue(dataList.get(i));
            }

            workbook.write(fos);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void exportXlsx_APPOTA(JComboBox<String> filePrefixField, List<List<String>> tidInfoList) {
        String srcFile = System.getProperty("user.dir") + File.separator + "sourceDocs" + File.separator
                + "UY-QUYEN-APPOTA.xlsx";
        String dstFile = System.getProperty("user.dir") + File.separator + "resultDocs" + File.separator
                + removeDiacritics(filePrefixField.getSelectedItem().toString().trim()).replace(" ", "-")
                + "_UY-QUYEN-APPOTA" + new SimpleDateFormat("_yyyy-MM-dd").format(new Date()) + ".xlsx";

        try (FileInputStream fis = new FileInputStream(srcFile);
             Workbook workbook = new XSSFWorkbook(fis);
             FileOutputStream fos = new FileOutputStream(dstFile)) {

            if (!tidInfoList.isEmpty()) {
                Sheet sheet = workbook.getSheet("Sheet1");
                for (int i = 0; i < tidInfoList.size(); i++) {
                    Row row = sheet.getRow(i + 2);
                    if (row == null) row = sheet.createRow(i + 2);

                    List<String> tidInfo = tidInfoList.get(i);
                    String[] cellValues = {
                            String.valueOf(i + 1), tidInfo.get(3), tidInfo.get(19), tidInfo.get(0),
                            tidInfo.get(1), tidInfo.get(2), tidInfo.get(6), tidInfo.get(5),
                            tidInfo.get(7), tidInfo.get(9), tidInfo.get(8), tidInfo.get(10)
                    };

                    for (int j = 0; j < cellValues.length; j++) {
                        Cell cell = row.getCell(j);
                        if (cell == null) cell = row.createCell(j);
                        cell.setCellValue(cellValues[j]);
                    }
                }
            }
            workbook.write(fos);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}