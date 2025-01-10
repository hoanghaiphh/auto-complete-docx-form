package utilities;

import org.apache.poi.xwpf.usermodel.*;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static utilities.StringConverter.removeDiacritics;

public class WordActions {

    @SafeVarargs
    public static void exportDocx(String fileName, JComboBox<String> filePrefixField,
                                  HashMap<String, String> replaceTexts, List<List<String>>... tidInfoLists) {
        String srcDocx = System.getProperty("user.dir") + File.separator + "sourceDocs" + File.separator
                + fileName + ".docx";
        String dstDocx = System.getProperty("user.dir") + File.separator + "resultDocs" + File.separator
                + removeDiacritics(filePrefixField.getSelectedItem().toString().trim()).replace(" ", "-")
                + "_" + fileName + new SimpleDateFormat("_yyyy-MM-dd").format(new Date()) + ".docx";

        try (FileInputStream fis = new FileInputStream(srcDocx);
             XWPFDocument document = new XWPFDocument(fis);
             FileOutputStream fos = new FileOutputStream(dstDocx)) {

            replaceInParagraphs(document, replaceTexts);
            replaceInTables(document, replaceTexts);

            if (tidInfoLists.length > 0 && tidInfoLists[0] != null) {
                List<List<String>> tidInfoList = tidInfoLists[0];

                if (srcDocx.contains("VINATTI")) {
                    if (srcDocx.contains("PHU-LUC")) {
                        insertDataIntoTable_VINATTI_PHULUC(document, tidInfoList);
                    } else if (srcDocx.contains("UY-QUYEN")) {
                        insertDataIntoTable_VINATTI_UYQUYEN(document, tidInfoList);
                    }

                } else if (srcDocx.contains("APPOTA")) {
                    insertDataIntoTable_APPOTA(document, tidInfoList);
                } else if (srcDocx.contains("ONEFIN")) {
                    insertDataIntoTable_ONEFIN(document, tidInfoList);
                }
            }

            document.write(fos);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void replaceInParagraphs(XWPFDocument document, HashMap<String, String> replaceTexts) {
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            for (XWPFRun run : paragraph.getRuns()) {
                for (Map.Entry<String, String> replaceText : replaceTexts.entrySet()) {
                    String text = run.getText(0);
                    if (text != null && text.contains(replaceText.getKey())) {
                        text = text.replace(replaceText.getKey(), replaceText.getValue());
                        run.setText(text, 0);
                    }
                }
            }
        }
    }

    private static void replaceInTables(XWPFDocument document, HashMap<String, String> replaceTexts) {
        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        for (XWPFRun run : paragraph.getRuns()) {
                            for (Map.Entry<String, String> replaceText : replaceTexts.entrySet()) {
                                String text = run.getText(0);
                                if (text != null && text.contains(replaceText.getKey())) {
                                    text = text.replace(replaceText.getKey(), replaceText.getValue());
                                    run.setText(text, 0);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    private static void insertDataIntoTable_ONEFIN(XWPFDocument document, List<List<String>> tidInfoList) {
        if (!tidInfoList.isEmpty()) {
            XWPFTable table = document.getTables().get(2);
            int startRow = 0;

            for (List<String> tidInfo : tidInfoList) {
                XWPFTableRow row = (startRow == 0) ? table.getRow(startRow++) : table.insertNewTableRow(startRow++);

                XWPFTableCell cell0 = row.getCell(0);
                if (cell0 == null) cell0 = row.createCell();
                insertParagraphsIntoCell(cell0, tidInfo.get(11));

                XWPFTableCell cell1 = row.getCell(1);
                if (cell1 == null) cell1 = row.createCell();
                insertParagraphsIntoCell(cell1, tidInfo.get(12));
            }
        }
    }

    private static void insertDataIntoTable_VINATTI_PHULUC(XWPFDocument document, List<List<String>> tidInfoList) {
        if (!tidInfoList.isEmpty()) {
            XWPFTable table = document.getTables().get(0);
            int startRow = 3;

            for (List<String> tidInfo : tidInfoList) {
                XWPFTableRow row;
                if (startRow == 3) {
                    row = table.getRow(startRow++);
                } else {
                    row = table.insertNewTableRow(startRow++);

                    XWPFTableCell cell0 = row.getCell(0);
                    if (cell0 == null) cell0 = row.createCell();
                    insertParagraphsIntoCell(cell0, "");
                }

                XWPFTableCell cell1 = row.getCell(1);
                if (cell1 == null) cell1 = row.createCell();
                insertParagraphsIntoCell(cell1, tidInfo.get(12), tidInfo.get(14), tidInfo.get(15), tidInfo.get(13));

                XWPFTableCell cell2 = row.getCell(2);
                if (cell2 == null) cell2 = row.createCell();
                insertParagraphsIntoCell(cell2, tidInfo.get(12), tidInfo.get(17), tidInfo.get(18), tidInfo.get(16));
            }
        }
    }

    private static void insertDataIntoTable_VINATTI_UYQUYEN(XWPFDocument document, List<List<String>> tidInfoList) {
        if (!tidInfoList.isEmpty()) {
            XWPFTable table = document.getTables().get(0);
            int startRow = 1;

            for (List<String> tidInfo : tidInfoList) {
                XWPFTableRow row = (startRow == 1) ? table.getRow(startRow) : table.insertNewTableRow(startRow);

                XWPFTableCell cell0 = row.getCell(0);
                if (cell0 == null) cell0 = row.createCell();
                insertParagraphsIntoCell(cell0, String.valueOf(startRow));

                XWPFTableCell cell1 = row.getCell(1);
                if (cell1 == null) cell1 = row.createCell();
                insertParagraphsIntoCell(cell1, tidInfo.get(1));

                XWPFTableCell cell2 = row.getCell(2);
                if (cell2 == null) cell2 = row.createCell();
                insertParagraphsIntoCell(cell2, tidInfo.get(8));

                XWPFTableCell cell3 = row.getCell(3);
                if (cell3 == null) cell3 = row.createCell();
                insertParagraphsIntoCell(cell3, tidInfo.get(9));

                XWPFTableCell cell4 = row.getCell(4);
                if (cell4 == null) cell4 = row.createCell();
                insertParagraphsIntoCell(cell4, tidInfo.get(10));

                XWPFTableCell cell5 = row.getCell(5);
                if (cell5 == null) cell5 = row.createCell();
                insertParagraphsIntoCell(cell5, "");

                startRow++;
            }
        }
    }

    private static void insertDataIntoTable_APPOTA(XWPFDocument document, List<List<String>> tidInfoList) {
        if (!tidInfoList.isEmpty()) {
            XWPFTable table = document.getTables().get(1);
            int startRow = 1;

            for (List<String> tidInfo : tidInfoList) {
                XWPFTableRow row = (startRow == 1) ? table.getRow(startRow) : table.insertNewTableRow(startRow);

                XWPFTableCell cell0 = row.getCell(0);
                if (cell0 == null) cell0 = row.createCell();
                insertParagraphsIntoCell(cell0, String.valueOf(startRow));

                XWPFTableCell cell1 = row.getCell(1);
                if (cell1 == null) cell1 = row.createCell();
                insertParagraphsIntoCell(cell1, tidInfo.get(4));

                XWPFTableCell cell2 = row.getCell(2);
                if (cell2 == null) cell2 = row.createCell();
                insertParagraphsIntoCell(cell2, tidInfo.get(2));

                XWPFTableCell cell3 = row.getCell(3);
                if (cell3 == null) cell3 = row.createCell();
                insertParagraphsIntoCell(cell3, tidInfo.get(6), tidInfo.get(5), tidInfo.get(7));

                XWPFTableCell cell4 = row.getCell(4);
                if (cell4 == null) cell4 = row.createCell();
                insertParagraphsIntoCell(cell4, tidInfo.get(9), tidInfo.get(8), tidInfo.get(10));

                XWPFTableCell cell5 = row.getCell(5);
                if (cell5 == null) cell5 = row.createCell();
                insertParagraphsIntoCell(cell5, "Giấy ủy quyền");

                startRow++;
            }
        }
    }

    private static void insertParagraphsIntoCell(XWPFTableCell cell, String... texts) {
        cell.removeParagraph(0);
        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        for (String text : texts) {
            XWPFParagraph paragraph = cell.addParagraph();
            paragraph.setSpacingAfter(0);
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun run = paragraph.createRun();
            run.setText(text);
            run.setFontFamily("Times New Roman");
            run.setFontSize(11);
        }
    }

}
