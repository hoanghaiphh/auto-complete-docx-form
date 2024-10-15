import org.apache.poi.xwpf.usermodel.*;

import javax.swing.*;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import java.awt.*;
import java.awt.event.FocusAdapter;
import java.awt.event.FocusEvent;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import static actions.NumberToWordsVN.convertToWords;

public class InputForm {
    private JPanel mainPanel;
    private JTabbedPane mainTabbedPane, lease;
    private JTextField authorizedName, authorizedId, authorizedIdDate, authorizedIdPlace,
            authorizedAddress, authorizedTel, authorizedAcc, authorizedEmail, authorizedBank,
            authorizerComName, authorizerComAddress, authorizerComIdDate, authorizerComIdPlace, contractNo,
            authorizerName, authorizerAddress, authorizerId, authorizerIdAddress, authorizerIdDate, authorizerIdPlace,
            quantity1, quantity2, quantity3, fee1, fee2, fee3, totalFee, totalFeeAsText, monthFee,
            handoverName, handoverId, handoverIdDate, handoverIdPlace;
    private JComboBox<String> authorizerComId, deviceName1, deviceName2, deviceName3, unitPrice1, unitPrice2, unitPrice3;
    private JButton exportButton;

    private HashMap<String, String> replaces = new HashMap<>();

    public InputForm() {
        Font newFont = new Font("Sogoe UI", Font.BOLD, 16);
        UIManager.put("TabbedPane.font", newFont);

        for (int i = 0; i < mainTabbedPane.getTabCount(); i++) {
            JLabel label = new JLabel(mainTabbedPane.getTitleAt(i));
            label.setFont(newFont);
            mainTabbedPane.setTabComponentAt(i, label);
        }

        for (int i = 0; i < lease.getTabCount(); i++) {
            JLabel label = new JLabel(lease.getTitleAt(i));
            label.setFont(newFont);
            lease.setTabComponentAt(i, label);
        }

        addListenerForInputField(deviceName1, unitPrice1, quantity1, fee1);
        addListenerForInputField(deviceName2, unitPrice2, quantity2, fee2);
        addListenerForInputField(deviceName3, unitPrice3, quantity3, fee3);

        DocumentListener documentListener = new DocumentListener() {
            public void insertUpdate(DocumentEvent e) { updateTotalFee(); }
            public void removeUpdate(DocumentEvent e) { updateTotalFee(); }
            public void changedUpdate(DocumentEvent e) { updateTotalFee(); }

            private void updateTotalFee() {
                int totalFeeValue = 0;
                String fee1Value = fee1.getText().replaceAll("\\.", "");
                if (!fee1Value.isEmpty()) {
                    totalFeeValue += Integer.parseInt(fee1Value);
                }
                String fee2Value = fee2.getText().replaceAll("\\.", "");
                if (!fee2Value.isEmpty()) {
                    totalFeeValue += Integer.parseInt(fee2Value);
                }
                String fee3Value = fee3.getText().replaceAll("\\.", "");
                if (!fee3Value.isEmpty()) {
                    totalFeeValue += Integer.parseInt(fee3Value);
                }

                NumberFormat numberFormat = NumberFormat.getInstance(new Locale("vi", "VN"));
                totalFee.setText(numberFormat.format(totalFeeValue));
                totalFeeAsText.setText(convertToWords(totalFeeValue) + " đồng");
            }
        };
        fee1.getDocument().addDocumentListener(documentListener);
        fee2.getDocument().addDocumentListener(documentListener);
        fee3.getDocument().addDocumentListener(documentListener);

        exportButton.addActionListener(e -> {
            getInputData();

            replaceTextInWord("TO-TRINH-CHO-THUE");
            replaceTextInWord("HOP-DONG-THUE");
            replaceTextInWord("BIEN-BAN-BAN-GIAO");
            replaceTextInWord("UY-QUYEN");
            replaceTextInWord("HOP-DONG-GIAO-KHOAN");

            JOptionPane.showMessageDialog(mainPanel, "Hồ sơ được xuất thành công!");
        });
    }

    public static void main(String[] args) {
        JFrame mainFrame = new JFrame("Hồ sơ thông tin");
        mainFrame.setContentPane(new InputForm().mainPanel);
        mainFrame.pack();
        mainFrame.setLocationRelativeTo(null);
        mainFrame.setResizable(false);
        mainFrame.setVisible(true);
        mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    private void getInputData() {
        replaces.clear();

        replaces.put("{authorizedName}", authorizedName.getText().trim().toUpperCase());
        replaces.put("{authorizedAddress}", authorizedAddress.getText().trim());
        replaces.put("{authorizedId}", authorizedId.getText().trim());
        replaces.put("{authorizedIdDate}", authorizedIdDate.getText().trim());
        replaces.put("{authorizedIdPlace}", authorizedIdPlace.getText().trim());
        replaces.put("{authorizedTel}", authorizedTel.getText().trim());
        replaces.put("{authorizedAcc}", authorizedAcc.getText().trim());
        replaces.put("{authorizedEmail}", authorizedEmail.getText().trim());
        replaces.put("{authorizedBank}", authorizedBank.getText().trim());

        replaces.put("{authorizerComName}", authorizerComName.getText().trim().toUpperCase());
        replaces.put("{authorizerComAddress}", authorizerComAddress.getText().trim());
        replaces.put("{authorizerComId}", authorizerComId.getSelectedItem().toString().trim());
        replaces.put("{authorizerComIdDate}", authorizerComIdDate.getText().trim());
        replaces.put("{authorizerComIdPlace}", authorizerComIdPlace.getText().trim());
        replaces.put("{contractNo}", contractNo.getText().trim());

        replaces.put("{authorizerName}", authorizerName.getText().trim().toUpperCase());
        replaces.put("{authorizerAddress}", authorizerAddress.getText().trim());
        replaces.put("{authorizerId}", authorizerId.getText().trim());
        replaces.put("{authorizerIdAddress}", authorizerIdAddress.getText().trim());
        replaces.put("{authorizerIdDate}", authorizerIdDate.getText().trim());
        replaces.put("{authorizerIdPlace}", authorizerIdPlace.getText().trim());

        String deviceName1Value = deviceName1.getSelectedItem().toString().trim();
        replaces.put("{deviceName1}", deviceName1Value.isEmpty() ? "" : ("Cho thuê máy POS " + deviceName1Value));
        String deviceName2Value = deviceName2.getSelectedItem().toString().trim();
        replaces.put("{deviceName2}", deviceName2Value.isEmpty() ? "" : ("Cho thuê máy POS " + deviceName2Value));
        replaces.put("{index2}", deviceName2Value.isEmpty() ? "" : "2");
        String deviceName3Value = deviceName3.getSelectedItem().toString().trim();
        replaces.put("{deviceName3}", deviceName3Value.isEmpty() ? "" : ("Cho thuê máy POS " + deviceName3Value));
        replaces.put("{index3}", deviceName3Value.isEmpty() ? "" : "3");

        replaces.put("{quantity1}", quantity1.getText().trim());
        replaces.put("{quantity2}", quantity2.getText().trim());
        replaces.put("{quantity3}", quantity3.getText().trim());
        replaces.put("{unitPrice1}", unitPrice1.getSelectedItem().toString().trim());
        replaces.put("{unitPrice2}", unitPrice2.getSelectedItem().toString().trim());
        replaces.put("{unitPrice3}", unitPrice3.getSelectedItem().toString().trim());
        replaces.put("{fee1}", fee1.getText().trim());
        replaces.put("{fee2}", fee2.getText().trim());
        replaces.put("{fee3}", fee3.getText().trim());
        replaces.put("{totalFee}", totalFee.getText().trim());
        replaces.put("{totalFeeAsText}", totalFeeAsText.getText().trim());
        replaces.put("{monthFee}", monthFee.getText().trim());

        replaces.put("{handoverName}", handoverName.getText().trim().toUpperCase());
        replaces.put("{handoverId}", handoverId.getText().trim());
        replaces.put("{handoverIdDate}", handoverIdDate.getText().trim());
        replaces.put("{handoverIdPlace}", handoverIdPlace.getText().trim());
    }

    private void replaceTextInWord(String fileName) {
        String inputFilePath = System.getProperty("user.dir") + File.separator + "sourceDocs" + File.separator
                + fileName + ".docx";
        String timestamp = new SimpleDateFormat("_yyyy-MM-dd").format(new Date());
        String outputFilePath = System.getProperty("user.dir") + File.separator + "resultDocs" + File.separator
                + authorizedId.getText() + "_" + fileName + timestamp + ".docx";

        FileInputStream fis = null;
        try {
            fis = new FileInputStream(inputFilePath);
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

            FileOutputStream fos = new FileOutputStream(outputFilePath);
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

    private void addCalculateListener(JComboBox unitPriceCB, JTextField quantityTF, JTextField feeTF) {
        String unitPriceCBValue = unitPriceCB.getSelectedItem().toString().replaceAll("\\.", "");
        String quantityTFValue = quantityTF.getText().replaceAll("\\.", "");

        if (!unitPriceCBValue.isEmpty() && !quantityTFValue.isEmpty()) {
            long quantity = Integer.parseInt(quantityTFValue);

            if (unitPriceCB.getSelectedIndex() != 0 && quantity != 0) {
                long unitPrice = Integer.parseInt(unitPriceCBValue);
                long fee = quantity * unitPrice;

                NumberFormat numberFormat = NumberFormat.getInstance(new Locale("vi", "VN"));
                feeTF.setText(numberFormat.format(fee));

            } else {
                feeTF.setText("");
            }

        } else {
            feeTF.setText("");
        }
    }

    private void addListenerForInputField(JComboBox deviceNameCB, JComboBox unitPriceCB, JTextField quantityTF, JTextField feeTF) {
        deviceNameCB.addActionListener(e -> {
            int select = deviceNameCB.getSelectedIndex();
            if (select == 1 || select == 2) {
                unitPriceCB.setSelectedIndex(select);
            }
        });

        unitPriceCB.getEditor().getEditorComponent().addKeyListener(new KeyAdapter() {
            @Override
            public void keyTyped(KeyEvent e) {
                char c = e.getKeyChar();
                if (!Character.isDigit(c) && c != '.') {
                    e.consume();
                }
            }
        });

        unitPriceCB.addActionListener(e -> {
            addCalculateListener(unitPriceCB, quantityTF, feeTF);
        });

        unitPriceCB.getEditor().getEditorComponent().addFocusListener(new FocusAdapter() {
            @Override
            public void focusLost(FocusEvent e) {
                addCalculateListener(unitPriceCB, quantityTF, feeTF);
            }
        });

        quantityTF.addKeyListener(new KeyAdapter() {
            @Override
            public void keyTyped(KeyEvent e) {
                char c = e.getKeyChar();
                if (!Character.isDigit(c) && c != '.') {
                    e.consume();
                }
            }
        });

        quantityTF.addActionListener(e -> {
            addCalculateListener(unitPriceCB, quantityTF, feeTF);
        });

        quantityTF.addFocusListener(new FocusAdapter() {
            @Override
            public void focusLost(FocusEvent e) {
                addCalculateListener(unitPriceCB, quantityTF, feeTF);
            }
        });
    }

}

