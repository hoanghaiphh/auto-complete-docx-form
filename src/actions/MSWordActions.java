package actions;

import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class MSWordActions {

    public static void replaceText(String srcDocx, String dstDocx, HashMap<String, String> replaces) {
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(srcDocx);
            XWPFDocument document = new XWPFDocument(fis);

            for (XWPFParagraph paragraph : document.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    for (Map.Entry<String, String> replace : replaces.entrySet()) {
                        String text = run.getText(0);
                        if (text != null && text.contains(replace.getKey())) {
                            text = text.replace(replace.getKey(), replace.getValue());
                        }
                        run.setText(text, 0);
                    }
                }
            }

            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            for (XWPFRun run : paragraph.getRuns()) {
                                for (Map.Entry<String, String> replace : replaces.entrySet()) {
                                    String text = run.getText(0);
                                    if (text != null && text.contains(replace.getKey())) {
                                        text = text.replace(replace.getKey(), replace.getValue());
                                    }
                                    run.setText(text, 0);
                                }
                            }
                        }
                    }
                }
            }

            FileOutputStream fos = new FileOutputStream(dstDocx);
            document.write(fos);

            fos.close();
            document.close();
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

    public static int countRowsInFirstTable(String fileName) {
        String filePath = System.getProperty("user.dir") + File.separator + "mergeDocs" + File.separator + fileName;
        int rowCount = 0;

        try (FileInputStream fis = new FileInputStream(filePath);
             XWPFDocument document = new XWPFDocument(fis)) {

            XWPFTable table = document.getTables().get(0);
            rowCount = table.getNumberOfRows();

        } catch (Exception e) {
            e.printStackTrace();
        }

        return rowCount;
    }

    public static void copyTableRow(String srcFileName, int srcRowIndex, int dstRowIndex) {
        String srcFile = System.getProperty("user.dir") + File.separator + "mergeDocs" + File.separator + srcFileName;
        String dstFile = System.getProperty("user.dir") + File.separator + "resultDocs" + File.separator
                + "UY-QUYEN-HOP-NHAT" + new SimpleDateFormat("_yyyy-MM-dd").format(new Date()) + ".docx";

        try (XWPFDocument srcDoc = new XWPFDocument(new FileInputStream(srcFile));
             XWPFDocument dstDoc = new XWPFDocument(new FileInputStream(dstFile))) {

            XWPFTable srcTable = srcDoc.getTables().get(0);
            XWPFTableRow srcRow = srcTable.getRow(srcRowIndex);

            XWPFTable dstTable = dstDoc.getTables().get(0);
            XWPFTableRow dstRow = dstTable.getRow(dstRowIndex);

            for (int i = 1; i <= 2; i++) {
                XWPFTableCell srcCell = srcRow.getCell(i);
                XWPFTableCell dstCell = dstRow.getCell(i);
                dstCell.removeParagraph(0);

                for (XWPFParagraph paragraph : srcCell.getParagraphs()) {
                    XWPFParagraph newParagraph = dstCell.addParagraph();
                    newParagraph.setSpacingAfter(0);
                    for (XWPFRun run : paragraph.getRuns()) {
                        XWPFRun newRun = newParagraph.createRun();
                        String text = run.getText(0);
                        if (text != null) {
                            newRun.setText(text);
                        }
                        newRun.setFontFamily("Times New Roman");
                        newRun.setFontSize(11);
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(dstFile)) {
                dstDoc.write(fos);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
