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
            quantity1, quantity2, quantity3, unitPrice1, unitPrice2, unitPrice3,
            fee1, fee2, fee3, totalFee, totalFeeAsText, monthFee, monthFee2, monthFee3,
            handoverName, handoverId, handoverIdDate, handoverIdPlace;
    private JComboBox<String> authorizerComId, deviceName1, deviceName2, deviceName3;
    private JRadioButton leaseRadio, sellRadio, sellWithTIDRadio;
    private JCheckBox addDevice2, addDevice3,
            to_trinh_thue, hop_dong_ban, bao_mat, bbbg, hop_dong_thue, uy_quyen, giao_khoan, to_trinh_ban;
    private JButton exportButton;

    public InputForm() {
        Font newFont = new Font("Sogoe UI Variable", Font.BOLD, 16);
        UIManager.put("TabbedPane.font", newFont);
        for (int i = 0; i < mainTabbedPane.getTabCount(); i++) {
            JLabel label = new JLabel(mainTabbedPane.getTitleAt(i));
            label.setFont(newFont);
            mainTabbedPane.setTabComponentAt(i, label);
        }

        authorizerComId.addItem("");
        for (String comId : getAllAuthorizerComIds()) {
            authorizerComId.addItem(comId);
        }

        addCalculationListener(deviceName1, unitPrice1, quantity1, fee1, monthFee);

        addCalculationListener(deviceName2, unitPrice2, quantity2, fee2, monthFee2);

        addCalculationListener(deviceName3, unitPrice3, quantity3, fee3, monthFee3);

        addDevice2.addActionListener(e -> {
            if (addDevice2.isSelected()) {
                deviceName2.setEnabled(true);
                quantity2.setEnabled(true);
                unitPrice2.setEnabled(true);
            } else {
                deviceName2.setEnabled(false);
                quantity2.setEnabled(false);
                unitPrice2.setEnabled(false);
                deviceName2.setSelectedIndex(0);
                quantity2.setText("");
                unitPrice2.setText("");
            }
        });

        addDevice3.addActionListener(e -> {
            if (addDevice3.isSelected()) {
                deviceName3.setEnabled(true);
                quantity3.setEnabled(true);
                unitPrice3.setEnabled(true);
            } else {
                deviceName3.setEnabled(false);
                quantity3.setEnabled(false);
                unitPrice3.setEnabled(false);
                deviceName3.setSelectedIndex(0);
                quantity3.setText("");
                unitPrice3.setText("");
            }
        });

        authorizedName.addActionListener(e -> {
            String authorizedNameValue = authorizedName.getText().trim();
            if (!authorizedNameValue.isEmpty()) {
                Row row = getRecord("KHACH_HANG", 1, authorizedNameValue);
                if (row != null) {
                    authorizedTel.setText(new DataFormatter().formatCellValue(row.getCell(2)));
                    authorizedEmail.setText(row.getCell(3).getStringCellValue());
                    authorizedId.setText(new DataFormatter().formatCellValue(row.getCell(4)));
                    authorizedIdDate.setText(row.getCell(5).getStringCellValue());
                    authorizedIdPlace.setText(row.getCell(6).getStringCellValue());
                    authorizedAddress.setText(row.getCell(7).getStringCellValue());
                    authorizedAcc.setText(new DataFormatter().formatCellValue(row.getCell(8)));
                    authorizedBank.setText(row.getCell(9).getStringCellValue());
                } else {
                    authorizedTel.setText("");
                    authorizedEmail.setText("");
                    authorizedId.setText("");
                    authorizedIdDate.setText("");
                    authorizedIdPlace.setText("");
                    authorizedAddress.setText("");
                    authorizedAcc.setText("");
                    authorizedBank.setText("");
                }
            }
        });

        authorizerComId.addActionListener(e -> {
            if (authorizerComId.getSelectedIndex() != 0) {
                String authorizerComIdValue = authorizerComId.getSelectedItem().toString().trim();
                Row row = getRecord("UY_QUYEN", 3, authorizerComIdValue);

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
            } else {
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
        });

        leaseRadio.addActionListener(e -> {
            if (leaseRadio.isSelected()) {
                to_trinh_thue.setSelected(true);
                hop_dong_ban.setSelected(false);
                bao_mat.setSelected(false);
                bbbg.setSelected(true);
                hop_dong_thue.setSelected(true);
                uy_quyen.setSelected(true);
                giao_khoan.setSelected(true);
                to_trinh_ban.setSelected(false);
            }
        });

        sellRadio.addActionListener(e -> {
            if (sellRadio.isSelected()) {
                to_trinh_thue.setSelected(false);
                hop_dong_ban.setSelected(true);
                bao_mat.setSelected(true);
                bbbg.setSelected(true);
                hop_dong_thue.setSelected(false);
                uy_quyen.setSelected(false);
                giao_khoan.setSelected(false);
                to_trinh_ban.setSelected(true);
            }
        });

        sellWithTIDRadio.addActionListener(e -> {
            if (sellWithTIDRadio.isSelected()) {
                to_trinh_thue.setSelected(false);
                hop_dong_ban.setSelected(true);
                bao_mat.setSelected(false);
                bbbg.setSelected(true);
                hop_dong_thue.setSelected(false);
                uy_quyen.setSelected(true);
                giao_khoan.setSelected(true);
                to_trinh_ban.setSelected(true);
            }
        });

        exportButton.addActionListener(e -> {
            HashMap<String, String> replaces = getDataInput();
            if (to_trinh_thue.isSelected()) {
                replaceTextInDocxFile("TO-TRINH-CHO-THUE", replaces);
            }
            if (hop_dong_ban.isSelected()) {
                replaceTextInDocxFile("HOP-DONG-BAN", replaces);
            }
            if (bao_mat.isSelected()) {
                replaceTextInDocxFile("THOA-THUAN-BAO-MAT", replaces);
            }
            if (hop_dong_thue.isSelected()) {
                replaceTextInDocxFile("HOP-DONG-THUE", replaces);
            }
            if (uy_quyen.isSelected()) {
                replaceTextInDocxFile("UY-QUYEN", replaces);
            }
            if (giao_khoan.isSelected()) {
                replaceTextInDocxFile("HOP-DONG-GIAO-KHOAN", replaces);
            }
            if (to_trinh_ban.isSelected()) {
                replaceTextInDocxFile("TO-TRINH-BAN", replaces);
            }
            replaceTextInDocxFile("BIEN-BAN-BAN-GIAO", replaces);

            if (!authorizedName.getText().trim().isEmpty()) {
                if (deviceName1.getSelectedIndex() != 0) {
                    List<String> customerInfo = saveCustomerInfo();
                    customerInfo.add(deviceName1.getSelectedItem().toString().trim());
                    addNewRecord("KHACH_HANG", customerInfo);
                }
                if (deviceName2.getSelectedIndex() != 0) {
                    List<String> customerInfo = saveCustomerInfo();
                    customerInfo.add(deviceName2.getSelectedItem().toString().trim());
                    addNewRecord("KHACH_HANG", customerInfo);
                }
                if (deviceName3.getSelectedIndex() != 0) {
                    List<String> customerInfo = saveCustomerInfo();
                    customerInfo.add(deviceName3.getSelectedItem().toString().trim());
                    addNewRecord("KHACH_HANG", customerInfo);
                }
            }

            if (authorizerComId.getSelectedIndex() == -1) {
                addNewRecord("UY_QUYEN", saveAuthorizerInfo());
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

    private HashMap<String, String> getDataInput() {
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
        replaces.put("{unitPrice1}", unitPrice1.getText().trim());
        replaces.put("{unitPrice2}", unitPrice2.getText().trim());
        replaces.put("{unitPrice3}", unitPrice3.getText().trim());
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

    private DocumentListener documentListener = new DocumentListener() {
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

    private void addCalculation(JTextField unitPrice, JTextField quantity, JTextField fee) {
        String unitPriceValue = unitPrice.getText().replaceAll("\\.", "");
        String quantityValue = quantity.getText().replaceAll("\\.", "");

        if (!unitPriceValue.isEmpty() && !quantityValue.isEmpty()) {
            int quantityInt = Integer.parseInt(quantityValue);
            int unitPriceInt = Integer.parseInt(unitPriceValue);

            int feeValue = quantityInt * unitPriceInt;
            NumberFormat numberFormat = NumberFormat.getInstance(new Locale("vi", "VN"));
            fee.setText(numberFormat.format(feeValue));

        } else {
            fee.setText("");
        }
    }

    private void addCalculationListener(JComboBox deviceName, JTextField unitPrice, JTextField quantity, JTextField fee, JTextField monthlyFee) {
        unitPrice.addKeyListener(new KeyAdapter() {
            @Override
            public void keyTyped(KeyEvent e) {
                char c = e.getKeyChar();
                if (!Character.isDigit(c) && c != '.') {
                    e.consume();
                }
            }
        });

        quantity.addKeyListener(new KeyAdapter() {
            @Override
            public void keyTyped(KeyEvent e) {
                char c = e.getKeyChar();
                if (!Character.isDigit(c) && c != '.') {
                    e.consume();
                }
            }
        });

        monthlyFee.addKeyListener(new KeyAdapter() {
            @Override
            public void keyTyped(KeyEvent e) {
                char c = e.getKeyChar();
                if (!Character.isDigit(c) && c != '.') {
                    e.consume();
                }
            }
        });

        unitPrice.addFocusListener(new FocusAdapter() {
            @Override
            public void focusLost(FocusEvent e) {
                addCalculation(unitPrice, quantity, fee);
            }
        });

        quantity.addFocusListener(new FocusAdapter() {
            @Override
            public void focusLost(FocusEvent e) {
                addCalculation(unitPrice, quantity, fee);
            }
        });

        unitPrice.addActionListener(e -> {
            addCalculation(unitPrice, quantity, fee);
        });

        quantity.addActionListener(e -> {
            addCalculation(unitPrice, quantity, fee);
        });

        fee.getDocument().addDocumentListener(documentListener);

        deviceName.addActionListener(e -> {
            int select = deviceName.getSelectedIndex();
            if (select == 1) {
                unitPrice.setText("10.000.000");
                monthlyFee.setText("350.000");
                addCalculation(unitPrice, quantity, fee);
            } else if (select == 2) {
                unitPrice.setText("5.000.000");
                monthlyFee.setText("250.000");
                addCalculation(unitPrice, quantity, fee);
            } else {
                unitPrice.setText("");
                monthlyFee.setText("");
                fee.setText("");
            }
        });
    }

    private List<String> saveAuthorizerInfo() {
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

    private List<String> saveCustomerInfo() {
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
        authorizedInfo.add("index 10");
        return authorizedInfo;
    }

}

