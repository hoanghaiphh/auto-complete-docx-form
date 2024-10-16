import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;

import javax.swing.*;
import javax.swing.event.*;
import java.awt.*;
import java.awt.event.FocusAdapter;
import java.awt.event.FocusEvent;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.io.File;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;

import static actions.NumberToWordsVN.convertToWords;
import static actions.MSExcelActions.*;
import static actions.MSWordActions.*;

public class InputForm {
    private JPanel mainPanel;
    private JTabbedPane mainTabbedPane;
    private JTextField authorizedName, authorizedId, authorizedIdDate, authorizedIdPlace,
            authorizedAddress, authorizedTel, authorizedAcc, authorizedEmail, authorizedBank,
            authorizerComName, authorizerComAddress, authorizerComIdDate, authorizerComIdPlace,
            authorizerName, authorizerAddress, authorizerId, authorizerIdAddress, authorizerIdDate, authorizerIdPlace,
            quantity1, quantity2, quantity3, fee1, fee2, fee3, totalFee, totalFeeAsText, monthFee,
            handoverName, handoverId, handoverIdDate, handoverIdPlace;
    private JComboBox<String> authorizerComId, deviceName1, deviceName2, deviceName3, unitPrice1, unitPrice2, unitPrice3;
    private JButton exportButton;

    private boolean isInitialized = false;
    public InputForm() {
        Font newFont = new Font("Sogoe UI", Font.BOLD, 16);
        UIManager.put("TabbedPane.font", newFont);
        for (int i = 0; i < mainTabbedPane.getTabCount(); i++) {
            JLabel label = new JLabel(mainTabbedPane.getTitleAt(i));
            label.setFont(newFont);
            mainTabbedPane.setTabComponentAt(i, label);
        }

        addListenerForInputField(deviceName1, unitPrice1, quantity1, fee1);
        addListenerForInputField(deviceName2, unitPrice2, quantity2, fee2);
        addListenerForInputField(deviceName3, unitPrice3, quantity3, fee3);

        DocumentListener documentListener = new DocumentListener() {
            public void insertUpdate(DocumentEvent e) {
                updateTotalFee();
            }
            public void removeUpdate(DocumentEvent e) {
                updateTotalFee();
            }
            public void changedUpdate(DocumentEvent e) {
                updateTotalFee();
            }
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

        authorizerComId.addPopupMenuListener(new PopupMenuListener() {
            @Override
            public void popupMenuWillBecomeVisible(PopupMenuEvent e) {
                if (!isInitialized) {
                    authorizerComId.addItem("");
                    for (String comId : getAllAuthorizerComIds()) {
                        authorizerComId.addItem(comId);
                    }
                    isInitialized = true;
                }
            }
            @Override
            public void popupMenuWillBecomeInvisible(PopupMenuEvent e) {
            }
            @Override
            public void popupMenuCanceled(PopupMenuEvent e) {
            }
        });

        authorizerComId.addActionListener(e -> {
            if (authorizerComId.getSelectedIndex() != 0) {
                String authorizerComIdValue = authorizerComId.getSelectedItem().toString().trim();
                Row row = getRowByAuthorizerComId(authorizerComIdValue);
                if (row != null) {
                    authorizerComName.setText(row.getCell(1).getStringCellValue());
                    authorizerComAddress.setText(row.getCell(2).getStringCellValue());
                    authorizerComIdDate.setText(row.getCell(4).getStringCellValue());
                    authorizerComIdPlace.setText(row.getCell(5).getStringCellValue());
                    authorizerName.setText(row.getCell(6).getStringCellValue());
                    authorizerIdAddress.setText(row.getCell(7).getStringCellValue());
                    authorizerAddress.setText(row.getCell(8).getStringCellValue());
                    authorizerId.setText(new DataFormatter().formatCellValue(row.getCell(9)));
                    authorizerIdDate.setText(row.getCell(10).getStringCellValue());
                    authorizerIdPlace.setText(row.getCell(11).getStringCellValue());
                } else {
                    clearAuthorizerInfo();
                }
            } else {
                clearAuthorizerInfo();
            }
        });

        exportButton.addActionListener(e -> {
            HashMap<String, String> replaces = getInputData();
            replaceTextInDocxFile("TO-TRINH-CHO-THUE", replaces);
            replaceTextInDocxFile("HOP-DONG-THUE", replaces);
            replaceTextInDocxFile("BIEN-BAN-BAN-GIAO", replaces);
            replaceTextInDocxFile("UY-QUYEN", replaces);
            replaceTextInDocxFile("HOP-DONG-GIAO-KHOAN", replaces);

            if (deviceName1.getSelectedIndex() != 0) {
                addNewRecord("KHACH_HANG", getAuthorizedInfo(deviceName1));
            }
            if (deviceName2.getSelectedIndex() != 0) {
                addNewRecord("KHACH_HANG", getAuthorizedInfo(deviceName2));
            }
            if (deviceName3.getSelectedIndex() != 0) {
                addNewRecord("KHACH_HANG", getAuthorizedInfo(deviceName3));
            }

            if (authorizerComId.getSelectedIndex() == -1) {
                addNewRecord("UY_QUYEN", getAuthorizerInfo());
                authorizerComId.addItem(authorizerComId.getSelectedItem().toString().trim());
            }

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

    private HashMap<String, String> getInputData() {
        HashMap<String, String> replaces = new HashMap<>();

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
        replaces.put("{contractNo}", authorizerComId.getSelectedItem().toString().trim() + "/ONEFIN");

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

        return replaces;
    }

    private void replaceTextInDocxFile(String fileName, HashMap<String, String> replaces) {
        String srcDocx = System.getProperty("user.dir") + File.separator + "sourceDocs" + File.separator
                + fileName + ".docx";
        String dstDocx = System.getProperty("user.dir") + File.separator + "resultDocs" + File.separator
                + authorizedId.getText() + "_" + fileName + new SimpleDateFormat("_yyyy-MM-dd").format(new Date()) + ".docx";
        replaceText(srcDocx, dstDocx, replaces);
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

    private void clearAuthorizerInfo() {
        authorizerComName.setText("");
        authorizerComAddress.setText("");
        authorizerComIdDate.setText("");
        authorizerComIdPlace.setText("");
        authorizerName.setText("");
        authorizerIdAddress.setText("");
        authorizerAddress.setText("");
        authorizerId.setText("");
        authorizerIdDate.setText("");
        authorizerIdPlace.setText("");
    }

    private List<String> getAuthorizerInfo() {
        List<String> authorizerInfo = new ArrayList<>();
        authorizerInfo.add("index 0");
        authorizerInfo.add(authorizerComName.getText().trim().toUpperCase());
        authorizerInfo.add(authorizerComAddress.getText().trim());
        authorizerInfo.add(authorizerComId.getSelectedItem().toString().trim());
        authorizerInfo.add(authorizerComIdDate.getText().trim());
        authorizerInfo.add(authorizerComIdPlace.getText().trim());
        authorizerInfo.add(authorizerName.getText().trim().toUpperCase());
        authorizerInfo.add(authorizerIdAddress.getText().trim());
        authorizerInfo.add(authorizerAddress.getText().trim());
        authorizerInfo.add(authorizerId.getText().trim());
        authorizerInfo.add(authorizerIdDate.getText().trim());
        authorizerInfo.add(authorizerIdPlace.getText().trim());
        return authorizerInfo;
    }

    private List<String> getAuthorizedInfo(JComboBox deviceName) {
        List<String> authorizedInfo = new ArrayList<>();
        authorizedInfo.add("index 0");
        authorizedInfo.add(authorizedName.getText().trim().toUpperCase());
        authorizedInfo.add(authorizedTel.getText().trim());
        authorizedInfo.add(authorizedEmail.getText().trim());
        authorizedInfo.add(authorizedId.getText().trim());
        authorizedInfo.add(authorizedIdDate.getText().trim());
        authorizedInfo.add(authorizedIdPlace.getText().trim());
        authorizedInfo.add(authorizedAddress.getText().trim());
        authorizedInfo.add(authorizedAcc.getText().trim());
        authorizedInfo.add(authorizedBank.getText().trim());
        authorizedInfo.add("# dia chi giao hang");
        authorizedInfo.add(deviceName.getSelectedItem().toString().trim());
        return authorizedInfo;
    }
}

