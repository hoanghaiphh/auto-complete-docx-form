package actions;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
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
}
