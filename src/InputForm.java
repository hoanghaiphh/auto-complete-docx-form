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
import java.util.stream.Collectors;

import static actions.NumberToWordsVN.convertToWords;
import static actions.MSExcelActions.*;
import static actions.MSWordActions.*;

public class InputForm {
    private JPanel mainPanel;
    private JTabbedPane mainTabbedPane;
    private JTextField authorizedIdDate, authorizedIdPlace, authorizedAddress, authorizedTel, authorizedEmail, authorizedAcc, authorizedBank,
            quantity1, quantity2, quantity3, unitPrice1, unitPrice2, unitPrice3, fee1, fee2, fee3, totalFee, totalFeeAsText, monthFee, monthFee2, monthFee3,
            handoverName, handoverId, handoverIdDate, handoverIdPlace,
            authorizerComAddress, authorizerComIdDate, authorizerComIdPlace, authorizerName, authorizerAddress, authorizerId, authorizerIdAddress, authorizerIdDate, authorizerIdPlace,
            mid1, mid2, mid3, mid4, mid5, mid6, mid7, mid8, tid1, tid2, tid3, tid4, tid5, tid6, tid7, tid8;
    private JComboBox<String> authorizedName, authorizedId, deviceName1, deviceName2, deviceName3, authorizerComName, authorizerComId, exportSelect;
    private JCheckBox addDevice2, addDevice3, addMid2, addMid3, addMid4, addMid5, addMid6, addMid7, addMid8,
            to_trinh_thue, hop_dong_ban, bao_mat, bb_giao_nhan, hop_dong_thue, uy_quyen, hd_giao_khoan, to_trinh_ban;
    private JButton exportButton;

    private static boolean authorizedNameFlag = false;

    public InputForm() {

        //  MAIN TABS
        Font newFont = new Font("Cambria", Font.PLAIN, 16);
        UIManager.put("TabbedPane.font", newFont);
        for (int i = 0; i < mainTabbedPane.getTabCount(); i++) {
            JLabel label = new JLabel(mainTabbedPane.getTitleAt(i));
            label.setFont(newFont);
            mainTabbedPane.setTabComponentAt(i, label);
        }

        //  MA_SO_DOANH_NGHIEP
        authorizerComId.getEditor().getEditorComponent().setForeground(new Color(0, 104, 139));

        authorizerComId.addItem("");
        for (String comId : getListValueOfColumn("UY_QUYEN", 3)) {
            authorizerComId.addItem(comId);
        }

        authorizerComId.addActionListener(e -> {
            if (authorizerComId.getSelectedIndex() != 0) {
                String authorizerComIdValue = authorizerComId.getSelectedItem().toString().trim();
                Row row = getRowByColumnValue("UY_QUYEN", 3, authorizerComIdValue);
                if (row != null) {
                    authorizerComName.setSelectedItem(row.getCell(1).getStringCellValue());
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
                authorizerComName.setSelectedIndex(0);
                clearAuthorizerInfo();
            }
        });

        //  TEN_DOANH_NGHIEP
        authorizerComName.getEditor().getEditorComponent().setForeground(new Color(0, 104, 139));

        authorizerComName.addItem("");
        for (String comName : getListValueOfColumn("UY_QUYEN", 1)) {
            authorizerComName.addItem(comName);
        }

        authorizerComName.addActionListener(e -> {
            if (authorizerComName.getSelectedIndex() != 0) {
                String authorizerComNameValue = authorizerComName.getSelectedItem().toString().trim();
                Row row = getRowByColumnValue("UY_QUYEN", 1, authorizerComNameValue);
                if (row != null) {
                    authorizerComAddress.setText(row.getCell(2).getStringCellValue());
                    authorizerComId.setSelectedItem(row.getCell(3).getStringCellValue());
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
                authorizerComId.setSelectedIndex(0);
                clearAuthorizerInfo();
            }
        });

        //  TEN_KHACH_HANG
        authorizedName.getEditor().getEditorComponent().setForeground(new Color(0, 104, 139));

        List<String> authorizedNameList = getListValueOfColumn("KHACH_HANG", 1);
        List<String> upperCaseList = authorizedNameList.stream().map(String::toUpperCase).toList();
        Set<String> dataList = new HashSet<>(upperCaseList);

        JTextField authorizedNameTF = (JTextField) authorizedName.getEditor().getEditorComponent();

        authorizedNameTF.getDocument().addDocumentListener(new DocumentListener() {
            @Override
            public void insertUpdate(DocumentEvent e) {
                updateComboBox();
            }

            @Override
            public void removeUpdate(DocumentEvent e) {
                updateComboBox();
            }

            @Override
            public void changedUpdate(DocumentEvent e) {
                updateComboBox();
            }

            private void updateComboBox() {
                if (authorizedNameFlag) {
                    return;
                }
                authorizedNameFlag = true;

                SwingUtilities.invokeLater(() -> {
                    String input = authorizedNameTF.getText();
                    if (input.trim().length() >= 2) {
                        List<String> filteredItems = dataList.stream()
                                .filter(item -> item.toUpperCase().contains(input.toUpperCase())).toList();
                        if (!filteredItems.isEmpty()) {
                            authorizedName.setModel(new DefaultComboBoxModel<>(filteredItems.toArray(new String[0])));
                            authorizedNameTF.setText(input);
                            authorizedName.showPopup();
                        } else {
                            authorizedName.hidePopup();
                        }
                    } else {
                        authorizedName.hidePopup();
                    }

                    authorizedNameFlag = false;
                });
            }
        });

        authorizedNameTF.addKeyListener(new KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent e) {
                if (e.getKeyCode() == KeyEvent.VK_DOWN || e.getKeyCode() == KeyEvent.VK_UP) {
                    int itemCount = authorizedName.getItemCount();
                    if (itemCount == 1) {
                        authorizedName.setSelectedIndex(0);
                        authorizedNameTF.setText((String) authorizedName.getSelectedItem());
                    } else if (itemCount > 1) {
                        authorizedNameFlag = true;
                        SwingUtilities.invokeLater(() -> authorizedNameFlag = false);
                        authorizedName.dispatchEvent(e);
                    }
                }
            }
        });

        authorizedName.addActionListener(e -> {
            authorizedId.removeAllItems();

            String authorizedNameValue = authorizedName.getSelectedItem().toString().trim();
            if (!authorizedNameValue.isEmpty()) {
                List<Row> rows = getListRowByColumnValue("KHACH_HANG", 1, authorizedNameValue);
                Set<String> uniqueSet = rows.stream()
                        .map(row -> new DataFormatter().formatCellValue(row.getCell(4))).collect(Collectors.toSet());
                List<String> uniqueList = new ArrayList<>(uniqueSet);

                if (!uniqueList.isEmpty()) {
                    authorizedId.addItem("");
                    for (String i : uniqueList) {
                        authorizedId.addItem(i);
                    }
                    if (uniqueList.size() == 1) {
                        authorizedId.setSelectedIndex(1);
                    } else {
                        authorizedId.setSelectedItem("< chọn căn cước >");
                        clearCustomerInfo();
                    }
                } else {
                    authorizedId.setSelectedItem("");
                    clearCustomerInfo();
                }
            } else {
                authorizedId.setSelectedItem("");
                clearCustomerInfo();
            }
        });

        //  CAN_CUOC_KHACH_HANG
        authorizedId.getEditor().getEditorComponent().setForeground(new Color(107, 107, 107));

        authorizedId.addActionListener(e -> {
            if (authorizedId.getSelectedIndex() > 0) {
                String authorizedIdValue = authorizedId.getSelectedItem().toString().trim();
                Row row = getRowByColumnValue("KHACH_HANG", 4, authorizedIdValue);

                authorizedTel.setText(new DataFormatter().formatCellValue(row.getCell(2)));
                authorizedEmail.setText(row.getCell(3).getStringCellValue());
                authorizedIdDate.setText(row.getCell(5).getStringCellValue());
                authorizedIdPlace.setText(row.getCell(6).getStringCellValue());
                authorizedAddress.setText(row.getCell(7).getStringCellValue());
                authorizedAcc.setText(new DataFormatter().formatCellValue(row.getCell(8)));
                authorizedBank.setText(row.getCell(9).getStringCellValue());
            } else if (authorizedId.getSelectedIndex() == 0) {
                clearCustomerInfo();
            }
        });

        //  EXPORT
        exportSelect.addActionListener(e -> {
            if (exportSelect.getSelectedIndex() == 1) {
                to_trinh_thue.setSelected(true);
                hop_dong_ban.setSelected(false);
                bao_mat.setSelected(false);
                bb_giao_nhan.setSelected(true);
                hop_dong_thue.setSelected(true);
                uy_quyen.setSelected(true);
                hd_giao_khoan.setSelected(true);
                to_trinh_ban.setSelected(false);
            } else if (exportSelect.getSelectedIndex() == 2) {
                to_trinh_thue.setSelected(false);
                hop_dong_ban.setSelected(true);
                bao_mat.setSelected(true);
                bb_giao_nhan.setSelected(true);
                hop_dong_thue.setSelected(false);
                uy_quyen.setSelected(false);
                hd_giao_khoan.setSelected(false);
                to_trinh_ban.setSelected(true);
            } else if (exportSelect.getSelectedIndex() == 3) {
                to_trinh_thue.setSelected(false);
                hop_dong_ban.setSelected(true);
                bao_mat.setSelected(false);
                bb_giao_nhan.setSelected(true);
                hop_dong_thue.setSelected(false);
                uy_quyen.setSelected(true);
                hd_giao_khoan.setSelected(true);
                to_trinh_ban.setSelected(true);
            } else {
                to_trinh_thue.setSelected(false);
                hop_dong_ban.setSelected(false);
                bao_mat.setSelected(false);
                bb_giao_nhan.setSelected(false);
                hop_dong_thue.setSelected(false);
                uy_quyen.setSelected(false);
                hd_giao_khoan.setSelected(false);
                to_trinh_ban.setSelected(false);
            }
        });

        exportButton.addActionListener(e -> {
            StringBuilder msg = new StringBuilder();

            if (!to_trinh_thue.isSelected() && !hop_dong_ban.isSelected() && !bao_mat.isSelected() &&
                    !hop_dong_thue.isSelected() && !uy_quyen.isSelected() && !hd_giao_khoan.isSelected() &&
                    !to_trinh_ban.isSelected() && !bb_giao_nhan.isSelected()) {
                msg.append("Không có File nào được xuất ra!\n");
                msg.append("\n");
            } else {
                msg.append("Hồ sơ xuất ra gồm có:\n");
                HashMap<String, String> replaces = getDataInput();
                if (to_trinh_thue.isSelected()) {
                    replaceTextInDocxFile("TO-TRINH-CHO-THUE", replaces);
                    msg.append("- Tờ trình cho thuê máy\n");
                }
                if (hop_dong_ban.isSelected()) {
                    replaceTextInDocxFile("HOP-DONG-BAN", replaces);
                    msg.append("- Hợp đồng mua bán\n");
                }
                if (bao_mat.isSelected()) {
                    replaceTextInDocxFile("THOA-THUAN-BAO-MAT", replaces);
                    msg.append("- Thỏa thuận bảo mật\n");
                }
                if (hop_dong_thue.isSelected()) {
                    replaceTextInDocxFile("HOP-DONG-THUE", replaces);
                    msg.append("- Hợp đồng thuê máy\n");
                }
                if (uy_quyen.isSelected()) {
                    replaceTextInDocxFile("UY-QUYEN", replaces);
                    msg.append("- Giấy ủy quyền\n");
                }
                if (hd_giao_khoan.isSelected()) {
                    replaceTextInDocxFile("HOP-DONG-GIAO-KHOAN", replaces);
                    msg.append("- Hợp đồng giao khoán\n");
                }
                if (to_trinh_ban.isSelected()) {
                    replaceTextInDocxFile("TO-TRINH-BAN", replaces);
                    msg.append("- Tờ trình bán máy\n");
                }
                if (bb_giao_nhan.isSelected()) {
                    replaceTextInDocxFile("BIEN-BAN-BAN-GIAO", replaces);
                    msg.append("- Biên bản giao nhận\n");
                }
                msg.append("\n");
            }

            // TODO: review excel file (add Quantity column / remove columns / ...)
            if (!authorizedName.getSelectedItem().toString().trim().isEmpty()
                    && !authorizedId.getSelectedItem().toString().trim().isEmpty()) {

                if (deviceName1.getSelectedIndex() != 0
                        || deviceName2.getSelectedIndex() != 0
                        || deviceName3.getSelectedIndex() != 0) {

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

                    msg.append("Thông tin KHÁCH HÀNG đã được lưu lại!\n");
                    msg.append("\n");
                } // TODO: save when no device? not exist -> save / save with mocking device
            }

            if (authorizerComId.getSelectedIndex() == -1 && authorizerComName.getSelectedIndex() == -1) {
                addNewRecord("UY_QUYEN", saveAuthorizerInfo());
                authorizerComId.addItem(authorizerComId.getSelectedItem().toString().trim());
                authorizerComName.addItem(authorizerComName.getSelectedItem().toString().trim().toUpperCase());
                msg.append("Thông tin ỦY QUYỀN đã được lưu lại!\n");
                msg.append("\n");
            }

            JOptionPane.showMessageDialog(mainPanel, msg);
        });

        //  THUE/MUA_MAY
        setEnabledDevice(addDevice2, deviceName2, unitPrice2, quantity2);
        setEnabledDevice(addDevice3, deviceName3, unitPrice3, quantity3);

        addCalculationListener(deviceName1, unitPrice1, quantity1, fee1, monthFee);
        addCalculationListener(deviceName2, unitPrice2, quantity2, fee2, monthFee2);
        addCalculationListener(deviceName3, unitPrice3, quantity3, fee3, monthFee3);

        // MID/TID
        setEnabledMidTid(addMid2, mid2, tid2);
        setEnabledMidTid(addMid3, mid3, tid3);
        setEnabledMidTid(addMid4, mid4, tid4);
        setEnabledMidTid(addMid5, mid5, tid5);
        setEnabledMidTid(addMid6, mid6, tid6);
        setEnabledMidTid(addMid7, mid7, tid7);
        setEnabledMidTid(addMid8, mid8, tid8);
    }

    public static void main(String[] args) {
        JFrame mainFrame = new JFrame("NGUYỄN THỊ HOÀNG YẾN");
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

    private String getMidTidList() {
        HashMap<String, String> allValues = new HashMap<>();

        String mid1Value = mid1.getText().trim();
        if (!mid1Value.isEmpty()) {
            allValues.put(mid1Value, tid1.getText().trim());
        }
        String mid2Value = mid2.getText().trim();
        if (!mid2Value.isEmpty()) {
            allValues.put(mid2Value, tid2.getText().trim());
        }
        String mid3Value = mid3.getText().trim();
        if (!mid3Value.isEmpty()) {
            allValues.put(mid3Value, tid3.getText().trim());
        }
        String mid4Value = mid4.getText().trim();
        if (!mid4Value.isEmpty()) {
            allValues.put(mid4Value, tid4.getText().trim());
        }
        String mid5Value = mid5.getText().trim();
        if (!mid5Value.isEmpty()) {
            allValues.put(mid5Value, tid5.getText().trim());
        }
        String mid6Value = mid6.getText().trim();
        if (!mid6Value.isEmpty()) {
            allValues.put(mid6Value, tid6.getText().trim());
        }
        String mid7Value = mid7.getText().trim();
        if (!mid7Value.isEmpty()) {
            allValues.put(mid7Value, tid7.getText().trim());
        }
        String mid8Value = mid8.getText().trim();
        if (!mid8Value.isEmpty()) {
            allValues.put(mid8Value, tid8.getText().trim());
        }

        List<Map.Entry<String, String>> entryList = new ArrayList<>(allValues.entrySet());
        StringBuilder midTidList = new StringBuilder();

        if (entryList.size() > 1) {
            for (int i = 0; i < entryList.size() - 1; i++) {
                midTidList.append("MiD_").append(i + 1).append(": ").append(entryList.get(i).getKey()).append(" ".repeat(10))
                        .append("TiD_").append(i + 1).append(": ").append(entryList.get(i).getValue()).append(" ".repeat(100));
            }
            midTidList.append("MiD_").append(entryList.size()).append(": ").append(entryList.get(entryList.size() - 1).getKey()).append(" ".repeat(10))
                    .append("TiD_").append(entryList.size()).append(": ").append(entryList.get(entryList.size() - 1).getValue());
        } else if (entryList.size() == 1) {
            midTidList.append("MiD: ").append(entryList.get(0).getKey()).append(" ".repeat(10)).append("TiD: ").append(entryList.get(0).getValue());
        }

        return midTidList.toString();
    }

    private HashMap<String, String> getDataInput() {
        HashMap<String, String> replaces = new HashMap<>();

        replaces.put("{authorizedName}", authorizedName.getSelectedItem().toString().trim().toUpperCase());
        replaces.put("{authorizedAddress}", authorizedAddress.getText().trim());
        replaces.put("{authorizedId}", authorizedId.getSelectedItem().toString().trim());
        replaces.put("{authorizedIdDate}", authorizedIdDate.getText().trim());
        replaces.put("{authorizedIdPlace}", authorizedIdPlace.getText().trim());
        replaces.put("{authorizedTel}", authorizedTel.getText().trim());
        replaces.put("{authorizedAcc}", authorizedAcc.getText().trim());
        replaces.put("{authorizedEmail}", authorizedEmail.getText().trim());
        replaces.put("{authorizedBank}", authorizedBank.getText().trim());

        replaces.put("{authorizerComName}", authorizerComName.getSelectedItem().toString().trim().toUpperCase());
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

        replaces.put("{midTidList}", getMidTidList());

        return replaces;
    }

    private void replaceTextInDocxFile(String fileName, HashMap<String, String> replaces) {
        String srcDocx = System.getProperty("user.dir") + File.separator + "sourceDocs" + File.separator
                + fileName + ".docx";
        String dstDocx = System.getProperty("user.dir") + File.separator + "resultDocs" + File.separator
                + authorizedId.getSelectedItem().toString().trim() + "_" + fileName + new SimpleDateFormat("_yyyy-MM-dd").format(new Date()) + ".docx";
        replaceText(srcDocx, dstDocx, replaces);
    }

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

        fee.getDocument().addDocumentListener(new DocumentListener() {
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

                if (totalFeeValue != 0) {
                    NumberFormat numberFormat = NumberFormat.getInstance(new Locale("vi", "VN"));
                    totalFee.setText(numberFormat.format(totalFeeValue));
                    totalFeeAsText.setText(convertToWords(totalFeeValue) + " đồng");
                } else {
                    totalFee.setText("");
                    totalFeeAsText.setText("");
                }
            }
        });

        deviceName.addActionListener(e -> {
            int select = deviceName.getSelectedIndex();
            if (select == 1 || select == 3) {
                unitPrice.setText("10.000.000");
                monthlyFee.setText("350.000");
                addCalculation(unitPrice, quantity, fee);
            } else if (select == 2) {
                unitPrice.setText("5.000.000");
                monthlyFee.setText("250.000");
                addCalculation(unitPrice, quantity, fee);
            } else if (select == 0) {
                quantity.setText("");
                unitPrice.setText("");
                monthlyFee.setText("");
                fee.setText("");
            } else {
                unitPrice.setText("");
                monthlyFee.setText("");
                fee.setText("");
            }
        });
    }

    private void setEnabledDevice(JCheckBox add, JComboBox deviceName, JTextField unitPrice, JTextField quantity) {
        add.addActionListener(e -> {
            if (add.isSelected()) {
                deviceName.setEnabled(true);
                deviceName.setEditable(true);
                quantity.setEnabled(true);
                quantity.setEditable(true);
                unitPrice.setEnabled(true);
                unitPrice.setEditable(true);
            } else {
                deviceName.setEnabled(false);
                quantity.setEnabled(false);
                unitPrice.setEnabled(false);
                deviceName.setSelectedIndex(0);
                quantity.setText("");
                unitPrice.setText("");
            }
        });
    }

    private void setEnabledMidTid(JCheckBox add, JTextField mid, JTextField tid) {
        add.addActionListener(e -> {
            if (add.isSelected()) {
                mid.setEnabled(true);
                mid.setEditable(true);
                tid.setEnabled(true);
                tid.setEditable(true);
            } else {
                mid.setEnabled(false);
                tid.setEnabled(false);
                mid.setText("");
                tid.setText("");
            }
        });
    }

    private void clearAuthorizerInfo() {
        authorizerComAddress.setText("");
        authorizerComIdDate.setText("");
        authorizerComIdPlace.setText("");
        authorizerName.setText("");
        authorizerIdAddress.setText("");
        authorizerAddress.setText("");
        authorizerId.setText("");
        authorizerIdDate.setText("");
        authorizerIdPlace.setText("Cục cảnh sát QLHC về TTXH");
        authorizerIdPlace.setCaretPosition(0);
    }

    private void clearCustomerInfo() {
        authorizedTel.setText("");
        authorizedEmail.setText("@gmail.com");
        authorizedEmail.setCaretPosition(0);
        authorizedIdDate.setText("");
        authorizedIdPlace.setText("Cục cảnh sát QLHC về TTXH");
        authorizedIdPlace.setCaretPosition(0);
        authorizedAddress.setText("");
        authorizedAcc.setText("");
        authorizedBank.setText("");
    }

    private List<String> saveAuthorizerInfo() {
        List<String> authorizerInfo = new ArrayList<>();
        authorizerInfo.add("index 0");
        authorizerInfo.add(authorizerComName.getSelectedItem().toString().trim().toUpperCase());
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
        authorizedInfo.add(authorizedName.getSelectedItem().toString().trim().toUpperCase());
        authorizedInfo.add(authorizedTel.getText().trim());
        authorizedInfo.add(authorizedEmail.getText().trim());
        authorizedInfo.add(authorizedId.getSelectedItem().toString().trim());
        authorizedInfo.add(authorizedIdDate.getText().trim());
        authorizedInfo.add(authorizedIdPlace.getText().trim());
        authorizedInfo.add(authorizedAddress.getText().trim());
        authorizedInfo.add(authorizedAcc.getText().trim());
        authorizedInfo.add(authorizedBank.getText().trim());
        authorizedInfo.add("dc_giao_hang");
        return authorizedInfo;
    }

}

