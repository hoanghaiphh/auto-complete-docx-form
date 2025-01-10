import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;

import javax.swing.*;
import javax.swing.event.*;
import java.awt.*;
import java.awt.event.*;
import java.text.NumberFormat;
import java.util.*;
import java.util.List;

import static utilities.StringConverter.*;
import static utilities.ExcelActions.*;
import static utilities.WordActions.*;

public class InputForm {
    private JPanel mainPanel;
    private JTabbedPane mainTabbedPane, authorizerTabbedPane;
    private JButton exportButton;
    private JComboBox<String> exportSelect;
    private JCheckBox to_trinh_thue, to_trinh_ban, hop_dong_thue, hop_dong_ban, bb_giao_nhan,
            uy_quyen, hd_giao_khoan, bao_mat, cam_ket;

    private JComboBox<String> authorizedName, authorizedId, deviceName1, deviceName2, deviceName3, handoverName;
    private JTextField authorizedIdDate, authorizedIdPlace, authorizedBirthday, authorizedAddress,
            authorizedAcc, authorizedBank, authorizedTel, authorizedEmail,
            quantity1, quantity2, quantity3, unitPrice1, unitPrice2, unitPrice3, fee1, fee2, fee3, monthFee2, monthFee3,
            monthFee, totalFee, totalFeeAsText, handoverId, handoverIdDate, handoverIdPlace;
    private JCheckBox addDevice1, addDevice2, addDevice3;

    private JRadioButton onefin, vinatti, appota;
    private JComboBox<String> authorizerComName;
    private JTextField authorizerComAddress, authorizerComId, authorizerComIdDate, authorizerComIdPlace,
            authorizerTaxCode, shopName, contractNo,
            authorizerName, authorizerId, authorizerIdDate, authorizerIdPlace, authorizerTel, authorizerEmail;

    private JCheckBox addTid1, addTid2, addTid3, addTid4, addTid5;
    private JTextField mid1, mid2, mid3, mid4, mid5, tid1, tid2, tid3, tid4, tid5,
            serie1, serie2, serie3, serie4, serie5, posName1, posName2, posName3, posName4, posName5,
            oldAccNo1, oldAccNo2, oldAccNo3, oldAccNo4, oldAccNo5,
            newAccNo1, newAccNo2, newAccNo3, newAccNo4, newAccNo5,
            oldAccBank1, oldAccBank2, oldAccBank3, oldAccBank4, oldAccBank5,
            newAccBank1, newAccBank2, newAccBank3, newAccBank4, newAccBank5;
    private JComboBox<String> oldAccName1, oldAccName2, oldAccName3, oldAccName4, oldAccName5,
            newAccName1, newAccName2, newAccName3, newAccName4, newAccName5;

    //  @formatter:off
    private final List<JCheckBox>
            ADD_DEVICE_LIST     = List.of(addDevice1, addDevice2, addDevice3);
    private final List<JComboBox<String>>
            DEVICE_NAME_LIST    = List.of(deviceName1, deviceName2, deviceName3);
    private final List<JTextField>
            UNIT_PRICE_LIST     = List.of(unitPrice1, unitPrice2, unitPrice3),
            QUANTITY_LIST       = List.of(quantity1, quantity2, quantity3),
            FEE_LIST            = List.of(fee1, fee2, fee3),
            MONTH_FEE_LIST      = List.of(monthFee, monthFee2, monthFee3);

    private final List<JCheckBox>
            ADD_TID_LIST        = List.of(addTid1, addTid2, addTid3, addTid4, addTid5);
    private final List<JComboBox<String>>
            OLD_ACC_NAME_LIST   = List.of(oldAccName1, oldAccName2, oldAccName3, oldAccName4, oldAccName5),
            NEW_ACC_NAME_LIST   = List.of(newAccName1, newAccName2, newAccName3, newAccName4, newAccName5);
    private final List<JTextField>
            TID_LIST            = List.of(tid1, tid2, tid3, tid4, tid5),
            MID_LIST            = List.of(mid1, mid2, mid3, mid4, mid5),
            SERIE_LIST          = List.of(serie1, serie2, serie3, serie4, serie5),
            POS_NAME_LIST       = List.of(posName1, posName2, posName3, posName4, posName5),
            OLD_ACC_NO_LIST     = List.of(oldAccNo1, oldAccNo2, oldAccNo3, oldAccNo4, oldAccNo5),
            OLD_ACC_BANK_LIST   = List.of(oldAccBank1, oldAccBank2, oldAccBank3, oldAccBank4, oldAccBank5),
            NEW_ACC_NO_LIST     = List.of(newAccNo1, newAccNo2, newAccNo3, newAccNo4, newAccNo5),
            NEW_ACC_BANK_LIST   = List.of(newAccBank1, newAccBank2, newAccBank3, newAccBank4, newAccBank5);
    //  @formatter:on

    private static class Flags {
        boolean suggestionFlag = false;
        static boolean settingValueFlag = false;
        static String serviceType = "";
    }


    public static void main(String[] args) {
        JFrame mainFrame = new JFrame("POS INFORMATION - VERSION 6.0 - ©COPYRIGHT BY HAIPH");
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

    public InputForm() {
        configFormUI();
        configCustomerComponents();
        configDeviceComponents();
        configHandoverStaffComponents();
        configPaymentIntermediariesComponents();
        configAuthorizerComponents();
        configTidComponents();
        configExportComponents();
    }

    private void configFormUI() {
        Font newFont = new Font("Cambria", Font.BOLD, 16);

        UIManager.put("TabbedPane.font", newFont);
        for (int i = 0; i < mainTabbedPane.getTabCount(); i++) {
            JLabel label = new JLabel(mainTabbedPane.getTitleAt(i));
            label.setFont(newFont);
            mainTabbedPane.setTabComponentAt(i, label);
        }

        for (int i = 0; i < authorizerTabbedPane.getTabCount(); i++) {
            JLabel label = new JLabel(authorizerTabbedPane.getTitleAt(i));
            label.setFont(newFont);
            authorizerTabbedPane.setTabComponentAt(i, label);
        }

        Color blue = new Color(0, 104, 139);
        authorizedName.getEditor().getEditorComponent().setForeground(blue);
        handoverName.getEditor().getEditorComponent().setForeground(blue);
        authorizerComName.getEditor().getEditorComponent().setForeground(blue);
    }

    private void configCustomerComponents() {
        addSuggestionForComboBox(authorizedName, "KHACH_HANG");
        authorizedName.addActionListener(e -> {
            String authorizedNameValue = getComboBoxValue(authorizedName).toUpperCase();
            if (!authorizedNameValue.isEmpty()) {
                List<Row> rows = getListRowByColumnValue("KHACH_HANG", 1, authorizedNameValue);
                List<String> uniqueList = rows.stream()
                        .map(row -> getFormattedCellValue(row, 5))
                        .distinct().toList();

                if (!uniqueList.isEmpty()) {
                    authorizedId.removeAllItems();
                    authorizedId.addItem("");
                    uniqueList.forEach(authorizedId::addItem);

                    if (uniqueList.size() == 1) {
                        authorizedId.setSelectedIndex(1);
                    } else {
                        authorizedId.setSelectedItem("< chọn căn cước >");
                        authorizedTel.setText("");
                        authorizedEmail.setText("");
                        authorizedIdDate.setText("");
                        authorizedIdPlace.setText("Cục cảnh sát QLHC về TTXH");
                        authorizedIdPlace.setCaretPosition(0);
                        authorizedAddress.setText("");
                        authorizedAcc.setText("");
                        authorizedBank.setText("");
                        authorizedBirthday.setText("");
                    }
                }
            } else {
                authorizedId.setSelectedItem("");
                authorizedTel.setText("");
                authorizedEmail.setText("");
                authorizedIdDate.setText("");
                authorizedIdPlace.setText("Cục cảnh sát QLHC về TTXH");
                authorizedIdPlace.setCaretPosition(0);
                authorizedAddress.setText("");
                authorizedAcc.setText("");
                authorizedBank.setText("");
                authorizedBirthday.setText("");
            }

            for (JComboBox<String> newAccName : NEW_ACC_NAME_LIST) {
                if (newAccName.isEnabled()) {
                    setComboBoxValue(newAccName, authorizedNameValue);
                }
            }
        });

        authorizedId.addActionListener(e -> {
            if (authorizedId.getSelectedIndex() > 0) {
                String authorizedIdValue = getComboBoxValue(authorizedId);
                List<Row> rows = getListRowByColumnValue("KHACH_HANG", 5, authorizedIdValue);
                Row row = rows.get(rows.size() - 1);
                authorizedBirthday.setText(getFormattedCellValue(row, 2));
                authorizedTel.setText(getFormattedCellValue(row, 3));
                authorizedEmail.setText(getFormattedCellValue(row, 4));
                authorizedIdDate.setText(getFormattedCellValue(row, 6));
                authorizedIdPlace.setText(getFormattedCellValue(row, 7));
                authorizedAddress.setText(getFormattedCellValue(row, 8));
                authorizedAcc.setText(getFormattedCellValue(row, 9));
                authorizedBank.setText(getFormattedCellValue(row, 10));

            } else if (authorizedId.getSelectedIndex() == 0) {
                authorizedBirthday.setText("");
                authorizedTel.setText("");
                authorizedEmail.setText("");
                authorizedIdDate.setText("");
                authorizedIdPlace.setText("Cục cảnh sát QLHC về TTXH");
                authorizedIdPlace.setCaretPosition(0);
                authorizedAddress.setText("");
                authorizedAcc.setText("");
                authorizedBank.setText("");
            }
        });

        authorizedAcc.addFocusListener(new FocusAdapter() {
            @Override
            public void focusLost(FocusEvent e) {
                for (JTextField newAccNo : NEW_ACC_NO_LIST) {
                    if (newAccNo.isEnabled()) {
                        newAccNo.setText(authorizedAcc.getText().trim());
                    }
                }
            }
        });

        authorizedBank.addFocusListener(new FocusAdapter() {
            @Override
            public void focusLost(FocusEvent e) {
                for (JTextField newAccBank : NEW_ACC_BANK_LIST) {
                    if (newAccBank.isEnabled()) {
                        newAccBank.setText(authorizedBank.getText().trim());
                    }
                }
            }
        });
    }

    private void configDeviceComponents() {
        for (int i = 0; i < ADD_DEVICE_LIST.size(); i++) {
            JCheckBox addDevice = ADD_DEVICE_LIST.get(i);
            JComboBox<String> deviceName = DEVICE_NAME_LIST.get(i);
            JTextField unitPrice = UNIT_PRICE_LIST.get(i), quantity = QUANTITY_LIST.get(i), fee = FEE_LIST.get(i),
                    mFee = MONTH_FEE_LIST.get(i);

            addDevice.addActionListener(e -> {
                if (addDevice.isSelected()) {
                    deviceName.setEnabled(true);
                    deviceName.setEditable(true);
                    quantity.setEnabled(true);
                    quantity.setEditable(true);
                    unitPrice.setEnabled(true);
                    unitPrice.setEditable(true);
                } else {
                    deviceName.setEnabled(false);
                    deviceName.setSelectedIndex(0);
                    quantity.setEnabled(false);
                    quantity.setText("");
                    unitPrice.setEnabled(false);
                    unitPrice.setText("");
                }
            });

            deviceName.addActionListener(e -> {
                int select = deviceName.getSelectedIndex();
                if (select == 1 || select == 3 || select == 4) {
                    unitPrice.setText("10.000.000");
                    mFee.setText("350.000");
                    setCalculationRule(unitPrice, quantity, fee);
                } else if (select == 2) {
                    unitPrice.setText("5.000.000");
                    mFee.setText("250.000");
                    setCalculationRule(unitPrice, quantity, fee);
                } else if (select == 0) {
                    quantity.setText("");
                    unitPrice.setText("");
                    mFee.setText("");
                    fee.setText("");
                } else {
                    unitPrice.setText("");
                    mFee.setText("");
                    fee.setText("");
                }
            });

            unitPrice.addActionListener(e -> setCalculationRule(unitPrice, quantity, fee));
            quantity.addActionListener(e -> setCalculationRule(unitPrice, quantity, fee));

            KeyAdapter keyAdapter = new KeyAdapter() {
                @Override
                public void keyTyped(KeyEvent e) {
                    char c = e.getKeyChar();
                    if (!Character.isDigit(c) && c != '.') {
                        e.consume();
                    }
                }
            };
            unitPrice.addKeyListener(keyAdapter);
            quantity.addKeyListener(keyAdapter);
            mFee.addKeyListener(keyAdapter);

            FocusAdapter focusAdapter = new FocusAdapter() {
                @Override
                public void focusLost(FocusEvent e) {
                    setCalculationRule(unitPrice, quantity, fee);
                }
            };
            unitPrice.addFocusListener(focusAdapter);
            quantity.addFocusListener(focusAdapter);

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
                        totalFeeAsText.setText(convertNumbersToWords(totalFeeValue) + " đồng");
                    } else {
                        totalFee.setText("");
                        totalFeeAsText.setText("");
                    }
                }
            });
        }
    }

    private void configHandoverStaffComponents() {
        addSuggestionForComboBox(handoverName, "NHAN_VIEN");
        handoverName.addActionListener(e -> {
            String handoverNameValue = getComboBoxValue(handoverName);
            if (!handoverNameValue.isEmpty()) {
                Row row = getRowByColumnValue("NHAN_VIEN", 1, handoverNameValue);
                if (row != null) {
                    handoverId.setText(getFormattedCellValue(row, 2));
                    handoverIdDate.setText(getFormattedCellValue(row, 3));
                    handoverIdPlace.setText(getFormattedCellValue(row, 4));
                }
            } else {
                handoverId.setText("");
                handoverIdDate.setText("");
                handoverIdPlace.setText("");
            }
        });
    }

    private void configPaymentIntermediariesComponents() {

        onefin.addActionListener(e -> {
            if (onefin.isSelected()) {
                authorizerComIdDate.setEnabled(true);
                authorizerComIdPlace.setEnabled(true);

                authorizerId.setEnabled(true);
                authorizerIdDate.setEnabled(true);
                authorizerIdPlace.setEnabled(true);

                contractNo.setEnabled(false);
                authorizerTaxCode.setEnabled(false);
                authorizerTel.setEnabled(false);
                authorizerEmail.setEnabled(false);
                shopName.setEnabled(false);

                for (int i = 0; i < ADD_TID_LIST.size(); i++) {
                    if (ADD_TID_LIST.get(i).isSelected()) {
                        setEnabledTidONEFIN(
                                MID_LIST.get(i), TID_LIST.get(i), SERIE_LIST.get(i), POS_NAME_LIST.get(i),
                                OLD_ACC_NAME_LIST.get(i), OLD_ACC_NO_LIST.get(i), OLD_ACC_BANK_LIST.get(i),
                                NEW_ACC_NAME_LIST.get(i), NEW_ACC_NO_LIST.get(i), NEW_ACC_BANK_LIST.get(i));
                    }
                }
            }
        });

        appota.addActionListener(e -> {
            if (appota.isSelected()) {
                authorizerComIdDate.setEnabled(false);
                authorizerComIdPlace.setEnabled(false);

                authorizerId.setEnabled(false);
                authorizerIdDate.setEnabled(false);
                authorizerIdPlace.setEnabled(false);

                contractNo.setEnabled(false);
                authorizerTaxCode.setEnabled(false);
                authorizerTel.setEnabled(false);
                authorizerEmail.setEnabled(false);
                shopName.setEnabled(true);

                for (int i = 0; i < ADD_TID_LIST.size(); i++) {
                    if (ADD_TID_LIST.get(i).isSelected()) {
                        setEnabledTidAPPOTA(
                                MID_LIST.get(i), TID_LIST.get(i), SERIE_LIST.get(i), POS_NAME_LIST.get(i),
                                OLD_ACC_NAME_LIST.get(i), OLD_ACC_NO_LIST.get(i), OLD_ACC_BANK_LIST.get(i),
                                NEW_ACC_NAME_LIST.get(i), NEW_ACC_NO_LIST.get(i), NEW_ACC_BANK_LIST.get(i));
                    }
                }

                addDefaultTidAccountWhenSwitchRadio();
            }
        });

        vinatti.addActionListener(e -> {
            if (vinatti.isSelected()) {
                authorizerComIdDate.setEnabled(false);
                authorizerComIdPlace.setEnabled(false);

                authorizerId.setEnabled(false);
                authorizerIdDate.setEnabled(false);
                authorizerIdPlace.setEnabled(false);

                contractNo.setEnabled(true);
                authorizerTaxCode.setEnabled(true);
                authorizerTel.setEnabled(true);
                authorizerEmail.setEnabled(true);
                shopName.setEnabled(false);

                for (int i = 0; i < ADD_TID_LIST.size(); i++) {
                    if (ADD_TID_LIST.get(i).isSelected()) {
                        setEnabledTidVINATTI(
                                MID_LIST.get(i), TID_LIST.get(i), SERIE_LIST.get(i), POS_NAME_LIST.get(i),
                                OLD_ACC_NAME_LIST.get(i), OLD_ACC_NO_LIST.get(i), OLD_ACC_BANK_LIST.get(i),
                                NEW_ACC_NAME_LIST.get(i), NEW_ACC_NO_LIST.get(i), NEW_ACC_BANK_LIST.get(i));
                    }
                }

                addDefaultTidAccountWhenSwitchRadio();
            }
        });

    }

    private void configAuthorizerComponents() {
        addSuggestionForComboBox(authorizerComName, "UY_QUYEN");
        authorizerComName.addActionListener(e -> {
            String authorizerComNameValue = getComboBoxValue(authorizerComName);
            if (!authorizerComNameValue.isEmpty()) {
                Row row = getRowByColumnValue("UY_QUYEN", 1, authorizerComNameValue);
                if (row != null) {
                    authorizerComAddress.setText(getFormattedCellValue(row, 2));
                    authorizerComId.setText(getFormattedCellValue(row, 3));
                    authorizerComIdDate.setText(getFormattedCellValue(row, 4));
                    authorizerComIdPlace.setText(getFormattedCellValue(row, 5));
                    authorizerTaxCode.setText(getFormattedCellValue(row, 6));
                    shopName.setText(getFormattedCellValue(row, 7));
                    authorizerName.setText(getFormattedCellValue(row, 8));
                    authorizerId.setText(getFormattedCellValue(row, 11));
                    authorizerIdDate.setText(getFormattedCellValue(row, 12));
                    authorizerIdPlace.setText(getFormattedCellValue(row, 13));
                    authorizerTel.setText(getFormattedCellValue(row, 14));
                    authorizerEmail.setText(getFormattedCellValue(row, 15));
                    contractNo.setText(getFormattedCellValue(row, 16));
                }
            } else {
                authorizerComId.setText("");
                authorizerComAddress.setText("");
                authorizerComIdDate.setText("");
                authorizerComIdPlace.setText("");
                authorizerName.setText("");
                authorizerId.setText("");
                authorizerIdDate.setText("");
                authorizerIdPlace.setText("Cục cảnh sát QLHC về TTXH");
                authorizerIdPlace.setCaretPosition(0);
                contractNo.setText("");
                authorizerTaxCode.setText("");
                authorizerTel.setText("");
                authorizerEmail.setText("");
                shopName.setText("");
            }
        });
    }

    private void configTidComponents() {
        for (int i = 0; i < ADD_TID_LIST.size(); i++) {
            JCheckBox addTid = ADD_TID_LIST.get(i);
            JComboBox<String> oldAccName = OLD_ACC_NAME_LIST.get(i), newAccName = NEW_ACC_NAME_LIST.get(i);
            JTextField mid = MID_LIST.get(i), tid = TID_LIST.get(i),
                    serie = SERIE_LIST.get(i), posName = POS_NAME_LIST.get(i),
                    oldAccNo = OLD_ACC_NO_LIST.get(i), oldAccBank = OLD_ACC_BANK_LIST.get(i),
                    newAccNo = NEW_ACC_NO_LIST.get(i), newAccBank = NEW_ACC_BANK_LIST.get(i);

            addTid.addActionListener(e -> {
                if (addTid.isSelected()) {
                    if (onefin.isSelected()) {
                        setEnabledTidONEFIN(mid, tid, serie, posName,
                                oldAccName, oldAccNo, oldAccBank, newAccName, newAccNo, newAccBank);

                    } else if (appota.isSelected()) {
                        setEnabledTidAPPOTA(mid, tid, serie, posName,
                                oldAccName, oldAccNo, oldAccBank, newAccName, newAccNo, newAccBank);

                        setComboBoxValue(oldAccName, "NGUYỄN BÁ BA");
                        setComboBoxValue(newAccName, getComboBoxValue(authorizedName).toUpperCase());

                    } else if (vinatti.isSelected()) {
                        setEnabledTidVINATTI(mid, tid, serie, posName,
                                oldAccName, oldAccNo, oldAccBank, newAccName, newAccNo, newAccBank);

                        setComboBoxValue(oldAccName, "NGUYỄN BÁ BA");
                        setComboBoxValue(newAccName, getComboBoxValue(authorizedName).toUpperCase());

                    } else {
                        JOptionPane.showMessageDialog(mainPanel, "VUI LÒNG CHỌN LOẠI ỦY QUYỀN !!!");
                    }
                } else {
                    mid.setEnabled(false);
                    mid.setText("");
                    tid.setEnabled(false);
                    tid.setText("");
                    serie.setEnabled(false);
                    serie.setText("");
                    posName.setEnabled(false);
                    posName.setText("");
                    oldAccName.setEnabled(false);
                    oldAccName.setSelectedItem("");
                    oldAccNo.setEnabled(false);
                    oldAccNo.setText("");
                    oldAccBank.setEnabled(false);
                    oldAccBank.setText("");
                    newAccName.setEnabled(false);
                    newAccName.setSelectedItem("");
                    newAccNo.setEnabled(false);
                    newAccNo.setText("");
                    newAccBank.setEnabled(false);
                    newAccBank.setText("");
                }
            });
        }

        List<JComboBox<String>> allAccNameList = new ArrayList<>();
        List<JTextField> allAccNoList = new ArrayList<>();
        List<JTextField> allAccBankList = new ArrayList<>();

        allAccNameList.addAll(OLD_ACC_NAME_LIST);
        allAccNameList.addAll(NEW_ACC_NAME_LIST);
        allAccNoList.addAll(OLD_ACC_NO_LIST);
        allAccNoList.addAll(NEW_ACC_NO_LIST);
        allAccBankList.addAll(OLD_ACC_BANK_LIST);
        allAccBankList.addAll(NEW_ACC_BANK_LIST);

        for (int i = 0; i < allAccNameList.size(); i++) {
            JComboBox<String> accName = allAccNameList.get(i);
            JTextField accNo = allAccNoList.get(i);
            JTextField accBank = allAccBankList.get(i);

            addSuggestionForComboBox(accName, "KHACH_HANG");

            accName.addActionListener(e -> {
                String accNameValue = getComboBoxValue(accName);
                if (!accNameValue.isEmpty()) {
                    List<Row> rows = getListRowByColumnValue("KHACH_HANG", 1, accNameValue);
                    if (!rows.isEmpty()) {
                        Row row = rows.get(rows.size() - 1);
                        accNo.setText(getFormattedCellValue(row, 9));
                        accBank.setText(getFormattedCellValue(row, 10));
                    } else if (accNameValue.equalsIgnoreCase(getComboBoxValue(authorizedName))) {
                        accNo.setText(authorizedAcc.getText().trim());
                        accBank.setText(authorizedBank.getText().trim());
                    }
                } else {
                    accNo.setText("");
                    accBank.setText("");
                }
            });
        }
    }

    private void configExportComponents() {

        exportSelect.addActionListener(e -> {
            if (exportSelect.getSelectedIndex() == 1) {
                to_trinh_thue.setSelected(true);
                to_trinh_ban.setSelected(false);
                hop_dong_thue.setSelected(true);
                hop_dong_ban.setSelected(false);
                hd_giao_khoan.setSelected(true);
                uy_quyen.setSelected(true);
                bao_mat.setSelected(false);
                bb_giao_nhan.setSelected(true);
            } else if (exportSelect.getSelectedIndex() == 2) {
                to_trinh_thue.setSelected(false);
                to_trinh_ban.setSelected(true);
                hop_dong_thue.setSelected(false);
                hop_dong_ban.setSelected(true);
                hd_giao_khoan.setSelected(false);
                uy_quyen.setSelected(false);
                bao_mat.setSelected(true);
                bb_giao_nhan.setSelected(true);
            } else if (exportSelect.getSelectedIndex() == 3) {
                to_trinh_thue.setSelected(false);
                to_trinh_ban.setSelected(true);
                hop_dong_thue.setSelected(false);
                hop_dong_ban.setSelected(true);
                hd_giao_khoan.setSelected(true);
                uy_quyen.setSelected(true);
                bao_mat.setSelected(true);
                bb_giao_nhan.setSelected(true);
            } else {
                to_trinh_thue.setSelected(false);
                to_trinh_ban.setSelected(false);
                hop_dong_thue.setSelected(false);
                hop_dong_ban.setSelected(false);
                hd_giao_khoan.setSelected(false);
                uy_quyen.setSelected(false);
                bao_mat.setSelected(false);
                bb_giao_nhan.setSelected(false);

                cam_ket.setSelected(false);
            }
        });

        exportButton.addActionListener(e -> {
            StringBuilder msg = new StringBuilder();

            //  Export Files
            if (!to_trinh_thue.isSelected() && !to_trinh_ban.isSelected() && !hop_dong_thue.isSelected()
                    && !hop_dong_ban.isSelected() && !hd_giao_khoan.isSelected() && !uy_quyen.isSelected()
                    && !bao_mat.isSelected() && !bb_giao_nhan.isSelected() && !cam_ket.isSelected()) {
                msg.append("Không có File nào được xuất ra!\n");
                msg.append("\n");
            } else {
                msg.append("Hồ sơ xuất ra gồm có:\n");
                HashMap<String, String> replaceTexts = getInputData();

                if (to_trinh_thue.isSelected()) {
                    exportDocx("TO-TRINH-CHO-THUE", authorizedName, replaceTexts);
                    msg.append("- Tờ trình cho thuê máy\n");
                }
                if (to_trinh_ban.isSelected()) {
                    exportDocx("TO-TRINH-BAN", authorizedName, replaceTexts);
                    msg.append("- Tờ trình bán máy\n");
                }
                if (hop_dong_thue.isSelected()) {
                    Flags.serviceType = "Cọc thuê máy POS ";
                    exportDocx("HOP-DONG-THUE", authorizedName, getInputData());
                    msg.append("- Hợp đồng thuê máy\n");
                }
                if (hop_dong_ban.isSelected()) {
                    Flags.serviceType = "Bán máy POS ";
                    exportDocx("HOP-DONG-BAN", authorizedName, getInputData());
                    msg.append("- Hợp đồng mua bán\n");
                }
                if (hd_giao_khoan.isSelected()) {
                    exportDocx("HOP-DONG-GIAO-KHOAN", authorizedName, replaceTexts);
                    msg.append("- Hợp đồng giao khoán\n");
                }
                if (uy_quyen.isSelected()) {
                    List<List<String>> tidInfoList = getTidInfoList();
                    if (onefin.isSelected()) {
                        exportDocx("UY-QUYEN-ONEFIN", authorizedName, replaceTexts, tidInfoList);
                        msg.append("- Ủy quyền ONEFIN\n");
                    } else if (appota.isSelected()) {
                        exportDocx("UY-QUYEN-APPOTA", authorizedName, replaceTexts, tidInfoList);
                        exportXlsx_APPOTA(authorizedName, tidInfoList);
                        msg.append("- Ủy quyền APPOTA (Word & Excel)\n");
                    } else if (vinatti.isSelected()) {
                        exportDocx("UY-QUYEN-VINATTI", authorizedName, replaceTexts, tidInfoList);
                        exportDocx("PHU-LUC-VINATTI", authorizedName, replaceTexts, tidInfoList);
                        msg.append("- Ủy quyền VINATTI (Ủy quyền + Phụ lục)\n");
                    } else {
                        JOptionPane.showMessageDialog(mainPanel, "VUI LÒNG CHỌN LOẠI ỦY QUYỀN !!!");
                        return;
                    }
                }
                if (bao_mat.isSelected()) {
                    exportDocx("THOA-THUAN-BAO-MAT", authorizedName, replaceTexts);
                    msg.append("- Thỏa thuận bảo mật\n");
                }
                if (bb_giao_nhan.isSelected()) {
                    exportDocx("BIEN-BAN-BAN-GIAO", authorizedName, replaceTexts);
                    msg.append("- Biên bản giao nhận\n");
                }
                if (cam_ket.isSelected()) {
                    exportDocx("CAM-KET", authorizedName, replaceTexts);
                    msg.append("- Giấy cam kết\n");
                }
                msg.append("\n");
            }

            //  Save Customer Info
            if (!getComboBoxValue(authorizedName).isEmpty() && !getComboBoxValue(authorizedId).isEmpty()) {
                if (deviceName1.getSelectedIndex() != 0 || deviceName2.getSelectedIndex() != 0
                        || deviceName3.getSelectedIndex() != 0) {

                    for (int i = 0; i < DEVICE_NAME_LIST.size(); i++) {
                        JComboBox<String> deviceName = DEVICE_NAME_LIST.get(i);
                        JTextField quantity = QUANTITY_LIST.get(i);

                        if (deviceName.getSelectedIndex() != 0) {
                            List<String> customerInfo = getCustomerInfo();
                            customerInfo.add(getComboBoxValue(deviceName));
                            customerInfo.add(quantity.getText().trim());
                            addNewRecord("KHACH_HANG", customerInfo);
                        }
                    }
                } else {
                    addNewRecord("KHACH_HANG", getCustomerInfo());
                }

                msg.append("Thông tin KHÁCH HÀNG đã được lưu lại!\n");
                msg.append("\n");
            }

            //  Save/Update Authorizer Info
            List<String> authorizerComNameList = getListValueOfColumn("UY_QUYEN", 1);
            String authorizerComNameValue = getComboBoxValue(authorizerComName).toUpperCase();
            if (!authorizerComNameValue.isEmpty()) {
                int index = authorizerComNameList.indexOf(authorizerComNameValue);

                if (index != -1) {
                    updateRecord("UY_QUYEN", index + 1, getAuthorizerInfo());
                } else {
                    addNewRecord("UY_QUYEN", getAuthorizerInfo());

                    msg.append("Thông tin ỦY QUYỀN đã được lưu lại!\n");
                    msg.append("\n");
                }
            }

            //  Save/Update Staff Info
            List<String> handoverNameList = getListValueOfColumn("NHAN_VIEN", 1);
            String handoverNameValue = getComboBoxValue(handoverName).toUpperCase();
            if (!handoverNameValue.isEmpty()) {
                int index = handoverNameList.indexOf(handoverNameValue);

                if (index != -1) {
                    updateRecord("NHAN_VIEN", index + 1, getStaffInfo());
                } else {
                    addNewRecord("NHAN_VIEN", getStaffInfo());

                    msg.append("Thông tin NHÂN VIÊN đã được lưu lại!\n");
                    msg.append("\n");
                }
            }

            JOptionPane.showMessageDialog(mainPanel, msg);
        });

    }


    private HashMap<String, String> getInputData() {
        HashMap<String, String> replaceTexts = new HashMap<>();

        replaceTexts.put("{authorizedName}", getComboBoxValue(authorizedName).toUpperCase());
        replaceTexts.put("{authorizedNameRD}", removeDiacritics(getComboBoxValue(authorizedName)));
        replaceTexts.put("{authorizedId}", getComboBoxValue(authorizedId));
        replaceTexts.put("{authorizedIdDate}", authorizedIdDate.getText().trim());
        replaceTexts.put("{authorizedIdPlace}", authorizedIdPlace.getText().trim());
        replaceTexts.put("{authorizedBirthday}", authorizedBirthday.getText().trim());
        replaceTexts.put("{authorizedAddress}", authorizedAddress.getText().trim());
        replaceTexts.put("{authorizedAcc}", authorizedAcc.getText().trim());
        replaceTexts.put("{authorizedBank}", authorizedBank.getText().trim());
        replaceTexts.put("{authorizedTel}", authorizedTel.getText().trim());
        replaceTexts.put("{authorizedEmail}", authorizedEmail.getText().trim());

        String deviceName1Value = getComboBoxValue(deviceName1);
        replaceTexts.put("{deviceName1}", deviceName1Value.isEmpty() ? "" : (Flags.serviceType + deviceName1Value));
        String deviceName2Value = getComboBoxValue(deviceName2);
        replaceTexts.put("{deviceName2}", deviceName2Value.isEmpty() ? "" : (Flags.serviceType + deviceName2Value));
        replaceTexts.put("{index2}", deviceName2Value.isEmpty() ? "" : "2");
        String deviceName3Value = getComboBoxValue(deviceName3);
        replaceTexts.put("{deviceName3}", deviceName3Value.isEmpty() ? "" : (Flags.serviceType + deviceName3Value));
        replaceTexts.put("{index3}", deviceName3Value.isEmpty() ? "" : "3");

        replaceTexts.put("{quantity1}", quantity1.getText().trim());
        replaceTexts.put("{quantity2}", quantity2.getText().trim());
        replaceTexts.put("{quantity3}", quantity3.getText().trim());
        replaceTexts.put("{unitPrice1}", unitPrice1.getText().trim());
        replaceTexts.put("{unitPrice2}", unitPrice2.getText().trim());
        replaceTexts.put("{unitPrice3}", unitPrice3.getText().trim());
        replaceTexts.put("{fee1}", fee1.getText().trim());
        replaceTexts.put("{fee2}", fee2.getText().trim());
        replaceTexts.put("{fee3}", fee3.getText().trim());
        replaceTexts.put("{totalFee}", totalFee.getText().trim());
        replaceTexts.put("{totalFeeAsText}", totalFeeAsText.getText().trim());
        replaceTexts.put("{monthFee}", monthFee.getText().trim());

        replaceTexts.put("{handoverName}", getComboBoxValue(handoverName).toUpperCase());
        replaceTexts.put("{handoverId}", handoverId.getText().trim());
        replaceTexts.put("{handoverIdDate}", handoverIdDate.getText().trim());
        replaceTexts.put("{handoverIdPlace}", handoverIdPlace.getText().trim());

        replaceTexts.put("{authorizerComName}", getComboBoxValue(authorizerComName).toUpperCase());
        replaceTexts.put("{authorizerComAddress}", authorizerComAddress.getText().trim());
        replaceTexts.put("{authorizerComId}", authorizerComId.getText().trim());
        replaceTexts.put("{authorizerComIdDate}", authorizerComIdDate.getText().trim());
        replaceTexts.put("{authorizerComIdPlace}", authorizerComIdPlace.getText().trim());
        replaceTexts.put("{authorizerTaxCode}", authorizerTaxCode.getText().trim());

        if (vinatti.isSelected()) {
            replaceTexts.put("{contractNo}", contractNo.getText().trim());
        } else if (onefin.isSelected()) {
            replaceTexts.put("{contractNo}", authorizerComId.getText().trim());
        }

        replaceTexts.put("{authorizerName}", authorizerName.getText().trim().toUpperCase());
        replaceTexts.put("{authorizerId}", authorizerId.getText().trim());
        replaceTexts.put("{authorizerIdDate}", authorizerIdDate.getText().trim());
        replaceTexts.put("{authorizerIdPlace}", authorizerIdPlace.getText().trim());
        replaceTexts.put("{authorizerTel}", authorizerTel.getText().trim());
        replaceTexts.put("{authorizerEmail}", authorizerEmail.getText().trim());

        return replaceTexts;
    }

    private List<List<String>> getTidInfoList() {
        List<List<String>> tidInfoList = new ArrayList<>();

        for (int i = 0; i < TID_LIST.size(); i++) {
            String tidValue = TID_LIST.get(i).getText().trim();
            if (!tidValue.isEmpty()) {
                List<String> tidInfo = new ArrayList<>();

                // 0
                tidInfo.add(MID_LIST.get(i).getText().trim().isEmpty() ? "" :
                        MID_LIST.get(i).getText().trim());
                // 1
                tidInfo.add(tidValue);
                // 2
                tidInfo.add(SERIE_LIST.get(i).getText().trim().isEmpty() ? "" :
                        SERIE_LIST.get(i).getText().trim());
                // 3
                tidInfo.add(POS_NAME_LIST.get(i).getText().trim().isEmpty() ? "" :
                        removeDiacritics(POS_NAME_LIST.get(i).getText().trim()));
                // 4
                tidInfo.add(shopName.getText().trim().isEmpty() ? "" :
                        shopName.getText().trim().toUpperCase());
                // 5
                tidInfo.add(OLD_ACC_NAME_LIST.get(i).getSelectedItem() == null ? "" :
                        removeDiacritics(getComboBoxValue(OLD_ACC_NAME_LIST.get(i))));
                // 6
                tidInfo.add(OLD_ACC_NO_LIST.get(i).getText().trim().isEmpty() ? "" :
                        OLD_ACC_NO_LIST.get(i).getText().trim());
                // 7
                tidInfo.add(OLD_ACC_BANK_LIST.get(i).getText().trim().isEmpty() ? "" :
                        OLD_ACC_BANK_LIST.get(i).getText().trim());
                // 8
                tidInfo.add(NEW_ACC_NAME_LIST.get(i).getSelectedItem() == null ? "" :
                        removeDiacritics(getComboBoxValue(NEW_ACC_NAME_LIST.get(i))));
                // 9
                tidInfo.add(NEW_ACC_NO_LIST.get(i).getText().trim().isEmpty() ? "" :
                        NEW_ACC_NO_LIST.get(i).getText().trim());
                // 10
                tidInfo.add(NEW_ACC_BANK_LIST.get(i).getText().trim().isEmpty() ? "" :
                        NEW_ACC_BANK_LIST.get(i).getText().trim());
                // ----------
                // 11
                tidInfo.add(MID_LIST.get(i).getText().trim().isEmpty() ? "" :
                        ("MID " + (i + 1) + ": " + MID_LIST.get(i).getText().trim()));
                // 12
                tidInfo.add("TID " + (i + 1) + ": " + tidValue);
                // 13
                tidInfo.add(OLD_ACC_NAME_LIST.get(i).getSelectedItem() == null ? "" :
                        ("Chủ TK: " + removeDiacritics(getComboBoxValue(OLD_ACC_NAME_LIST.get(i)))));
                // 14
                tidInfo.add(OLD_ACC_NO_LIST.get(i).getText().trim().isEmpty() ? "" :
                        ("STK: " + OLD_ACC_NO_LIST.get(i).getText().trim()));
                // 15
                tidInfo.add(OLD_ACC_BANK_LIST.get(i).getText().trim().isEmpty() ? "" :
                        ("Mở tại: " + OLD_ACC_BANK_LIST.get(i).getText().trim()));
                // 16
                tidInfo.add(NEW_ACC_NAME_LIST.get(i).getSelectedItem() == null ? "" :
                        ("Chủ TK: " + removeDiacritics(getComboBoxValue(NEW_ACC_NAME_LIST.get(i)))));
                // 17
                tidInfo.add(NEW_ACC_NO_LIST.get(i).getText().trim().isEmpty() ? "" :
                        ("STK: " + NEW_ACC_NO_LIST.get(i).getText().trim()));
                // 18
                tidInfo.add(NEW_ACC_BANK_LIST.get(i).getText().trim().isEmpty() ? "" :
                        ("Mở tại: " + NEW_ACC_BANK_LIST.get(i).getText().trim()));
                // ----------
                // 19
                tidInfo.add(authorizerComAddress.getText().trim().isEmpty() ? "" :
                        authorizerComAddress.getText().trim());

                tidInfoList.add(tidInfo);
            }
        }
        return tidInfoList;
    }

    private List<String> getCustomerInfo() {
        List<String> authorizedInfo = new ArrayList<>();
        authorizedInfo.add("index 0");
        authorizedInfo.add(getComboBoxValue(authorizedName).toUpperCase());
        authorizedInfo.add(authorizedBirthday.getText().trim());
        authorizedInfo.add(authorizedTel.getText().trim());
        authorizedInfo.add(authorizedEmail.getText().trim());
        authorizedInfo.add(getComboBoxValue(authorizedId));
        authorizedInfo.add(authorizedIdDate.getText().trim());
        authorizedInfo.add(authorizedIdPlace.getText().trim());
        authorizedInfo.add(authorizedAddress.getText().trim());
        authorizedInfo.add(authorizedAcc.getText().trim());
        authorizedInfo.add(authorizedBank.getText().trim());
        authorizedInfo.add("");
        return authorizedInfo;
    }

    private List<String> getStaffInfo() {
        List<String> staffInfo = new ArrayList<>();
        staffInfo.add("index 0");
        staffInfo.add(getComboBoxValue(handoverName).toUpperCase());
        staffInfo.add(handoverId.getText().trim());
        staffInfo.add(handoverIdDate.getText().trim());
        staffInfo.add(handoverIdPlace.getText().trim());
        return staffInfo;
    }

    private List<String> getAuthorizerInfo() {
        List<String> authorizerInfo = new ArrayList<>();
        authorizerInfo.add("index 0");
        authorizerInfo.add(getComboBoxValue(authorizerComName).toUpperCase());
        authorizerInfo.add(authorizerComAddress.getText().trim());
        authorizerInfo.add(authorizerComId.getText().trim());
        authorizerInfo.add(authorizerComIdDate.getText().trim());
        authorizerInfo.add(authorizerComIdPlace.getText().trim());
        authorizerInfo.add(authorizerTaxCode.getText().trim());
        authorizerInfo.add(shopName.getText().trim().toUpperCase());
        authorizerInfo.add(authorizerName.getText().trim().toUpperCase());
        authorizerInfo.add("");
        authorizerInfo.add("");
        authorizerInfo.add(authorizerId.getText().trim());
        authorizerInfo.add(authorizerIdDate.getText().trim());
        authorizerInfo.add(authorizerIdPlace.getText().trim());
        authorizerInfo.add(authorizerTel.getText().trim());
        authorizerInfo.add(authorizerEmail.getText().trim());
        authorizerInfo.add(contractNo.getText().trim());
        return authorizerInfo;
    }

    private String getComboBoxValue(JComboBox<String> comboBox) {
        return comboBox.getSelectedItem() != null ? comboBox.getSelectedItem().toString().trim() : "";
    }

    private String getFormattedCellValue(Row row, int cellIndex) {
        DataFormatter formatter = new DataFormatter();
        return row.getCell(cellIndex) != null ? formatter.formatCellValue(row.getCell(cellIndex)) : "";
    }

    private void addSuggestionForComboBox(JComboBox<String> comboBox, String sheetName) {
        Flags flags = new Flags();
        JTextField comboBoxTF = (JTextField) comboBox.getEditor().getEditorComponent();

        comboBoxTF.getDocument().addDocumentListener(new DocumentListener() {
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
                if (flags.suggestionFlag || Flags.settingValueFlag) {
                    return;
                }
                flags.suggestionFlag = true;

                List<String> comboBoxList = getListValueOfColumn(sheetName, 1);
                List<String> upperCaseList = comboBoxList.stream().map(String::toUpperCase).toList();
                Set<String> dataList = new HashSet<>(upperCaseList);

                SwingUtilities.invokeLater(() -> {
                    String input = comboBoxTF.getText();
                    if (input.trim().length() >= 2) {
                        List<String> filteredItems = dataList.stream()
                                .filter(item -> item.toUpperCase().contains(input.toUpperCase())).toList();
                        if (!filteredItems.isEmpty()) {
                            comboBox.setModel(new DefaultComboBoxModel<>(filteredItems.toArray(new String[0])));
                            comboBoxTF.setText(input);
                            if (comboBox.isShowing()) {
                                comboBox.showPopup();
                            }
                        } else {
                            comboBox.hidePopup();
                        }
                    } else {
                        comboBox.hidePopup();
                    }

                    flags.suggestionFlag = false;
                });
            }
        });

        comboBoxTF.addKeyListener(new KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent e) {
                if (e.getKeyCode() == KeyEvent.VK_DOWN || e.getKeyCode() == KeyEvent.VK_UP) {
                    int itemCount = comboBox.getItemCount();
                    if (itemCount == 1) {
                        comboBox.setSelectedIndex(0);
                        comboBoxTF.setText((String) comboBox.getSelectedItem());
                    } else if (itemCount > 1) {
                        flags.suggestionFlag = true;
                        SwingUtilities.invokeLater(() -> flags.suggestionFlag = false);
                        comboBox.dispatchEvent(e);
                    }
                }
            }
        });
    }

    private void setComboBoxValue(JComboBox<String> comboBox, String value) {
        Flags.settingValueFlag = true;
        comboBox.setSelectedItem(value);
        Flags.settingValueFlag = false;
    }

    private void setCalculationRule(JTextField unitPrice, JTextField quantity, JTextField fee) {
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

    private void addDefaultTidAccountWhenSwitchRadio() {
        for (int i = 0; i < OLD_ACC_NAME_LIST.size(); i++) {
            JComboBox<String> oldAccName = OLD_ACC_NAME_LIST.get(i);

            if (oldAccName.isEnabled() && getComboBoxValue(oldAccName).isEmpty()
                    && OLD_ACC_NO_LIST.get(i).getText().trim().isEmpty()
                    && OLD_ACC_BANK_LIST.get(i).getText().trim().isEmpty()) {

                setComboBoxValue(oldAccName, "NGUYỄN BÁ BA");
            }
        }

        for (int i = 0; i < NEW_ACC_NAME_LIST.size(); i++) {
            JComboBox<String> newAccName = NEW_ACC_NAME_LIST.get(i);

            if (newAccName.isEnabled() && getComboBoxValue(newAccName).isEmpty()
                    && NEW_ACC_NO_LIST.get(i).getText().trim().isEmpty()
                    && NEW_ACC_BANK_LIST.get(i).getText().trim().isEmpty()) {

                setComboBoxValue(newAccName, getComboBoxValue(authorizedName).toUpperCase());
            }
        }
    }

    private void setEnabledTidONEFIN(JTextField mid, JTextField tid, JTextField serie, JTextField posName,
                                     JComboBox<String> oldAccName, JTextField oldAccNo, JTextField oldAccBank,
                                     JComboBox<String> newAccName, JTextField newAccNo, JTextField newAccBank) {
        mid.setEnabled(true);
        mid.setEditable(true);
        tid.setEnabled(true);
        tid.setEditable(true);
        serie.setEnabled(false);
        serie.setText("");
        posName.setEnabled(false);
        posName.setText("");
        oldAccName.setEnabled(false);
        oldAccName.setSelectedItem("");
        oldAccNo.setEnabled(false);
        oldAccNo.setText("");
        oldAccBank.setEnabled(false);
        oldAccBank.setText("");
        newAccName.setEnabled(false);
        newAccName.setSelectedItem("");
        newAccNo.setEnabled(false);
        newAccNo.setText("");
        newAccBank.setEnabled(false);
        newAccBank.setText("");
    }

    private void setEnabledTidAPPOTA(JTextField mid, JTextField tid, JTextField serie, JTextField posName,
                                     JComboBox<String> oldAccName, JTextField oldAccNo, JTextField oldAccBank,
                                     JComboBox<String> newAccName, JTextField newAccNo, JTextField newAccBank) {
        mid.setEnabled(true);
        mid.setEditable(true);
        tid.setEnabled(true);
        tid.setEditable(true);
        serie.setEnabled(true);
        serie.setEditable(true);
        posName.setEnabled(true);
        posName.setEditable(true);
        oldAccName.setEnabled(true);
        oldAccName.setEditable(true);
        oldAccNo.setEnabled(true);
        oldAccNo.setEditable(true);
        oldAccBank.setEnabled(true);
        oldAccBank.setEditable(true);
        newAccName.setEnabled(true);
        newAccName.setEditable(true);
        newAccNo.setEnabled(true);
        newAccNo.setEditable(true);
        newAccBank.setEnabled(true);
        newAccBank.setEditable(true);
    }

    private void setEnabledTidVINATTI(JTextField mid, JTextField tid, JTextField serie, JTextField posName,
                                      JComboBox<String> oldAccName, JTextField oldAccNo, JTextField oldAccBank,
                                      JComboBox<String> newAccName, JTextField newAccNo, JTextField newAccBank) {
        mid.setEnabled(false);
        mid.setText("");
        tid.setEnabled(true);
        tid.setEditable(true);
        serie.setEnabled(false);
        serie.setText("");
        posName.setEnabled(false);
        posName.setText("");
        oldAccName.setEnabled(true);
        oldAccName.setEditable(true);
        oldAccNo.setEnabled(true);
        oldAccNo.setEditable(true);
        oldAccBank.setEnabled(true);
        oldAccBank.setEditable(true);
        newAccName.setEnabled(true);
        newAccName.setEditable(true);
        newAccNo.setEnabled(true);
        newAccNo.setEditable(true);
        newAccBank.setEnabled(true);
        newAccBank.setEditable(true);
    }

}