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
    private JTextField authorizedIdDate, authorizedIdPlace, authorizedAddress, authorizedTel, authorizedEmail, authorizedAcc,
            authorizedBank, quantity1, quantity2, quantity3, unitPrice1, unitPrice2, unitPrice3, fee1, fee2, fee3,
            totalFee, totalFeeAsText, monthFee, monthFee2, monthFee3, handoverId, handoverIdDate, handoverIdPlace,
            authorizerComAddress, authorizerComId, authorizerComIdDate, authorizerComIdPlace, authorizerName,
            authorizerId, authorizerIdDate, authorizerIdPlace, mid1, mid2, mid3, mid4, mid5, tid1, tid2, tid3, tid4, tid5,
            authorizedBirthday, shopName, vinattiNo, authorizerTaxCode, authorizerTel, authorizerEmail, serie, posName;
    private JComboBox<String> deviceName1, deviceName2, deviceName3,
            authorizedName, authorizedId, authorizerComName, exportSelect, handoverName;
    private JCheckBox addDevice2, addDevice3, addMid2, addMid3, addMid4, addMid5, to_trinh_thue, hop_dong_ban, bao_mat,
            bb_giao_nhan, hop_dong_thue, uy_quyen, hd_giao_khoan, to_trinh_ban, cam_ket;
    private JRadioButton onefin, vinatti, appota;
    private JButton exportButton;

    private static String serviceType = "", key1 = "", key2 = "", key3 = "", key4 = "", key0 = "",
            value1 = "", value2 = "", value3 = "", value4 = "", value0 = "";

    public InputForm() {
        //  MAIN TABS
        Font newFont = new Font("Cambria", Font.PLAIN, 16);
        UIManager.put("TabbedPane.font", newFont);
        for (int i = 0; i < mainTabbedPane.getTabCount(); i++) {
            JLabel label = new JLabel(mainTabbedPane.getTitleAt(i));
            label.setFont(newFont);
            mainTabbedPane.setTabComponentAt(i, label);
        }

        //  TEN_DOANH_NGHIEP
        ComboBoxFlags authorizerComNameFlag = new ComboBoxFlags();
        addSuggestionListener(authorizerComName, authorizerComNameFlag, "UY_QUYEN");
        authorizerComName.addActionListener(e -> {
            String authorizerComNameValue = authorizerComName.getSelectedItem().toString().trim();
            if (!authorizerComNameValue.isEmpty()) {
                Row row = getRowByColumnValue("UY_QUYEN", 1, authorizerComNameValue);
                if (row != null) {
                    authorizerComAddress.setText(row.getCell(2) != null ? row.getCell(2).getStringCellValue() : "");
                    authorizerComId.setText(row.getCell(3) != null ? row.getCell(3).getStringCellValue() : "");
                    authorizerComIdDate.setText(row.getCell(4) != null ? row.getCell(4).getStringCellValue() : "");
                    authorizerComIdPlace.setText(row.getCell(5) != null ? row.getCell(5).getStringCellValue() : "");
                    authorizerTaxCode.setText(row.getCell(6) != null ? new DataFormatter().formatCellValue(row.getCell(6)) : "");
                    shopName.setText(row.getCell(7) != null ? row.getCell(7).getStringCellValue() : "");
                    authorizerName.setText(row.getCell(8) != null ? row.getCell(8).getStringCellValue() : "");
                    authorizerId.setText(row.getCell(11) != null ? new DataFormatter().formatCellValue(row.getCell(11)) : "");
                    authorizerIdDate.setText(row.getCell(12) != null ? row.getCell(12).getStringCellValue() : "");
                    authorizerIdPlace.setText(row.getCell(13) != null ? row.getCell(13).getStringCellValue() : "");
                    authorizerTel.setText(row.getCell(14) != null ? new DataFormatter().formatCellValue(row.getCell(14)) : "");
                    authorizerEmail.setText(row.getCell(15) != null ? row.getCell(15).getStringCellValue() : "");
                    vinattiNo.setText(row.getCell(16) != null ? new DataFormatter().formatCellValue(row.getCell(16)) : "");
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
                vinattiNo.setText("");
                authorizerTaxCode.setText("");
                authorizerTel.setText("");
                authorizerEmail.setText("");
                shopName.setText("");
            }
        });

        //  TEN_KHACH_HANG
        ComboBoxFlags authorizedNameFlag = new ComboBoxFlags();
        addSuggestionListener(authorizedName, authorizedNameFlag, "KHACH_HANG");
        authorizedName.addActionListener(e -> {
            String authorizedNameValue = authorizedName.getSelectedItem().toString().trim();
            if (!authorizedNameValue.isEmpty()) {
                List<Row> rows = getListRowByColumnValue("KHACH_HANG", 1, authorizedNameValue);
                Set<String> uniqueSet = rows.stream()
                        .map(row -> new DataFormatter().formatCellValue(row.getCell(5))).collect(Collectors.toSet());
                List<String> uniqueList = new ArrayList<>(uniqueSet);

                if (!uniqueList.isEmpty()) {
                    authorizedId.removeAllItems();
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
                List<Row> rows = getListRowByColumnValue("KHACH_HANG", 5, authorizedIdValue);
                Row row = rows.get(rows.size() - 1);

                authorizedBirthday.setText(row.getCell(2) != null ? row.getCell(2).getStringCellValue() : "");
                authorizedTel.setText(row.getCell(3) != null ? new DataFormatter().formatCellValue(row.getCell(3)) : "");
                authorizedEmail.setText(row.getCell(4) != null ? row.getCell(4).getStringCellValue() : "");
                authorizedIdDate.setText(row.getCell(6) != null ? row.getCell(6).getStringCellValue() : "");
                authorizedIdPlace.setText(row.getCell(7) != null ? row.getCell(7).getStringCellValue() : "");
                authorizedAddress.setText(row.getCell(8) != null ? row.getCell(8).getStringCellValue() : "");
                authorizedAcc.setText(row.getCell(9) != null ? new DataFormatter().formatCellValue(row.getCell(9)) : "");
                authorizedBank.setText(row.getCell(10) != null ? row.getCell(10).getStringCellValue() : "");

            } else if (authorizedId.getSelectedIndex() == 0) {
                clearCustomerInfo();
            }
        });

        //  NGUOI_BAN_GIAO
        ComboBoxFlags handoverNameFlag = new ComboBoxFlags();
        addSuggestionListener(handoverName, handoverNameFlag, "NHAN_VIEN");
        handoverName.addActionListener(e -> {
            String handoverNameValue = handoverName.getSelectedItem().toString().trim();
            if (!handoverNameValue.isEmpty()) {
                Row row = getRowByColumnValue("NHAN_VIEN", 1, handoverNameValue);
                if (row != null) {
                    handoverId.setText(row.getCell(2) != null ? row.getCell(2).getStringCellValue() : "");
                    handoverIdDate.setText(row.getCell(3) != null ? row.getCell(3).getStringCellValue() : "");
                    handoverIdPlace.setText(row.getCell(4) != null ? row.getCell(4).getStringCellValue() : "");
                }
            } else {
                handoverId.setText("");
                handoverIdDate.setText("");
                handoverIdPlace.setText("");
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
                bao_mat.setSelected(true);
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
                cam_ket.setSelected(false);
            }
        });
        exportButton.addActionListener(e -> {
            StringBuilder msg = new StringBuilder();

            if (!to_trinh_thue.isSelected() && !hop_dong_ban.isSelected() && !bao_mat.isSelected() &&
                    !hop_dong_thue.isSelected() && !uy_quyen.isSelected() && !hd_giao_khoan.isSelected() &&
                    !to_trinh_ban.isSelected() && !bb_giao_nhan.isSelected() && !cam_ket.isSelected()) {
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
                    serviceType = "Bán máy POS ";
                    replaceTextInDocxFile("HOP-DONG-BAN", getDataInput());
                    msg.append("- Hợp đồng mua bán\n");
                }
                if (bao_mat.isSelected()) {
                    replaceTextInDocxFile("THOA-THUAN-BAO-MAT", replaces);
                    msg.append("- Thỏa thuận bảo mật\n");
                }
                if (hop_dong_thue.isSelected()) {
                    serviceType = "Cọc thuê máy POS ";
                    replaceTextInDocxFile("HOP-DONG-THUE", getDataInput());
                    msg.append("- Hợp đồng thuê máy\n");
                }
                if (uy_quyen.isSelected()) {
                    if (onefin.isSelected()) {
                        replaceTextInDocxFile("UY-QUYEN-ONEFIN", replaces);
                        msg.append("- Giấy ủy quyền ONEFIN\n");
                    } else if (appota.isSelected()) {
                        replaceTextInDocxFile("UY-QUYEN-APPOTA", replaces);
                        exportAppotaExcel();
                        msg.append("- Giấy ủy quyền APPOTA (Word & Excel)\n");
                    } else if (vinatti.isSelected()) {
                        var vinattiTidList = getVinattiTidList();
                        if (vinattiTidList.size() == 1) {
                            key0 = vinattiTidList.get(0).getKey();
                            value0 = vinattiTidList.get(0).getValue();
                            replaceTextInDocxFile("UY-QUYEN-VINATTI", getDataInput());
                        } else if (vinattiTidList.size() == 2) {
                            key0 = vinattiTidList.get(0).getKey();
                            value0 = vinattiTidList.get(0).getValue();
                            key1 = vinattiTidList.get(1).getKey();
                            value1 = vinattiTidList.get(1).getValue();
                            replaceTextInDocxFile("UY-QUYEN-VINATTI-2", getDataInput());
                        } else if (vinattiTidList.size() == 3) {
                            key0 = vinattiTidList.get(0).getKey();
                            value0 = vinattiTidList.get(0).getValue();
                            key1 = vinattiTidList.get(1).getKey();
                            value1 = vinattiTidList.get(1).getValue();
                            key2 = vinattiTidList.get(2).getKey();
                            value2 = vinattiTidList.get(2).getValue();
                            replaceTextInDocxFile("UY-QUYEN-VINATTI-3", getDataInput());
                        } else if (vinattiTidList.size() == 4) {
                            key0 = vinattiTidList.get(0).getKey();
                            value0 = vinattiTidList.get(0).getValue();
                            key1 = vinattiTidList.get(1).getKey();
                            value1 = vinattiTidList.get(1).getValue();
                            key2 = vinattiTidList.get(2).getKey();
                            value2 = vinattiTidList.get(2).getValue();
                            key3 = vinattiTidList.get(3).getKey();
                            value3 = vinattiTidList.get(3).getValue();
                            replaceTextInDocxFile("UY-QUYEN-VINATTI-4", getDataInput());
                        } else if (vinattiTidList.size() == 5) {
                            key0 = vinattiTidList.get(0).getKey();
                            value0 = vinattiTidList.get(0).getValue();
                            key1 = vinattiTidList.get(1).getKey();
                            value1 = vinattiTidList.get(1).getValue();
                            key2 = vinattiTidList.get(2).getKey();
                            value2 = vinattiTidList.get(2).getValue();
                            key3 = vinattiTidList.get(3).getKey();
                            value3 = vinattiTidList.get(3).getValue();
                            key4 = vinattiTidList.get(4).getKey();
                            value4 = vinattiTidList.get(4).getValue();
                            replaceTextInDocxFile("UY-QUYEN-VINATTI-5", getDataInput());
                        } else {
                            replaceTextInDocxFile("UY-QUYEN-VINATTI", replaces);
                        }
                        msg.append("- Giấy ủy quyền VINATTI\n");
                    } else {
                        JOptionPane.showMessageDialog(mainPanel, "VUI LÒNG CHỌN LOẠI ỦY QUYỀN !!!");
                        return;
                    }
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
                if (cam_ket.isSelected()) {
                    replaceTextInDocxFile("CAM-KET", replaces);
                    msg.append("- Giấy cam kết\n");
                }
                msg.append("\n");
            }

            if (!authorizedName.getSelectedItem().toString().trim().isEmpty()
                    && !authorizedId.getSelectedItem().toString().trim().isEmpty()) {

                if (deviceName1.getSelectedIndex() != 0
                        || deviceName2.getSelectedIndex() != 0
                        || deviceName3.getSelectedIndex() != 0) {

                    if (deviceName1.getSelectedIndex() != 0) {
                        List<String> customerInfo = saveCustomerInfo();
                        customerInfo.add(deviceName1.getSelectedItem().toString().trim());
                        customerInfo.add(quantity1.getText().trim());
                        addNewRecord("KHACH_HANG", customerInfo);
                    }
                    if (deviceName2.getSelectedIndex() != 0) {
                        List<String> customerInfo = saveCustomerInfo();
                        customerInfo.add(deviceName2.getSelectedItem().toString().trim());
                        customerInfo.add(quantity2.getText().trim());
                        addNewRecord("KHACH_HANG", customerInfo);
                    }
                    if (deviceName3.getSelectedIndex() != 0) {
                        List<String> customerInfo = saveCustomerInfo();
                        customerInfo.add(deviceName3.getSelectedItem().toString().trim());
                        customerInfo.add(quantity3.getText().trim());
                        addNewRecord("KHACH_HANG", customerInfo);
                    }
                } else {
                    addNewRecord("KHACH_HANG", saveCustomerInfo());
                }

                msg.append("Thông tin KHÁCH HÀNG đã được lưu lại!\n");
                msg.append("\n");
            }

            List<String> authorizerComNameList = getListValueOfColumn("UY_QUYEN", 1);
            String authorizerComNameValue = authorizerComName.getSelectedItem().toString().trim().toUpperCase();
            if (!authorizerComNameValue.isEmpty()) {
                int index = authorizerComNameList.indexOf(authorizerComNameValue);
                if (index != -1) {
                    updateRecord(FILE_NAME, FILE_NAME, "UY_QUYEN", index + 1, saveAuthorizerInfo());
                } else {
                    addNewRecord("UY_QUYEN", saveAuthorizerInfo());
                    msg.append("Thông tin ỦY QUYỀN đã được lưu lại!\n");
                    msg.append("\n");
                }
            }

            List<String> handoverNameList = getListValueOfColumn("NHAN_VIEN", 1);
            String handoverNameValue = handoverName.getSelectedItem().toString().trim().toUpperCase();
            if (!handoverNameValue.isEmpty()) {
                int index = handoverNameList.indexOf(handoverNameValue);
                if (index != -1) {
                    updateRecord(FILE_NAME, FILE_NAME, "NHAN_VIEN", index + 1, saveStaffInfo());
                } else {
                    addNewRecord("NHAN_VIEN", saveStaffInfo());
                    msg.append("Thông tin NHÂN VIÊN đã được lưu lại!\n");
                    msg.append("\n");
                }
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

        // TRUNG_GIAN_THANH_TOAN
        onefin.addActionListener(e -> {
            if (onefin.isSelected()) {
                authorizerComIdDate.setEnabled(true);
                authorizerComIdPlace.setEnabled(true);

                authorizerId.setEnabled(true);
                authorizerIdDate.setEnabled(true);
                authorizerIdPlace.setEnabled(true);

                vinattiNo.setEnabled(false);
                authorizerTaxCode.setEnabled(false);
                authorizerTel.setEnabled(false);
                authorizerEmail.setEnabled(false);
                shopName.setEnabled(false);

                serie.setText("");
                posName.setText("");
                serie.setEnabled(false);
                posName.setEnabled(false);
            }
        });
        appota.addActionListener(e -> {
            if (appota.isSelected()) {
                authorizerComIdDate.setEnabled(false);
                authorizerComIdPlace.setEnabled(false);

                authorizerId.setEnabled(false);
                authorizerIdDate.setEnabled(false);
                authorizerIdPlace.setEnabled(false);

                vinattiNo.setEnabled(false);
                authorizerTaxCode.setEnabled(false);
                authorizerTel.setEnabled(false);
                authorizerEmail.setEnabled(false);
                shopName.setEnabled(true);

                serie.setEnabled(true);
                serie.setEditable(true);
                posName.setEnabled(true);
                posName.setEditable(true);
            }
        });
        vinatti.addActionListener(e -> {
            if (vinatti.isSelected()) {
                authorizerComIdDate.setEnabled(false);
                authorizerComIdPlace.setEnabled(false);

                authorizerId.setEnabled(false);
                authorizerIdDate.setEnabled(false);
                authorizerIdPlace.setEnabled(false);

                vinattiNo.setEnabled(true);
                vinattiNo.setEditable(true);
                authorizerTaxCode.setEnabled(true);
                authorizerTel.setEnabled(true);
                authorizerEmail.setEnabled(true);
                shopName.setEnabled(false);

                serie.setText("");
                posName.setText("");
                serie.setEnabled(false);
                posName.setEnabled(false);
            }
        });
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

    private List<Map.Entry<String, String>> getMidTidList() {
        HashMap<String, String> allValues = new HashMap<>();

        String mid1Value = mid1.getText().trim();
        String tid1Value = tid1.getText().trim();
        if (!mid1Value.isEmpty() && !tid1Value.isEmpty()) {
            allValues.put(mid1Value, tid1Value);
        }
        String mid2Value = mid2.getText().trim();
        String tid2Value = tid2.getText().trim();
        if (!mid2Value.isEmpty() && !tid2Value.isEmpty()) {
            allValues.put(mid2Value, tid2Value);
        }
        String mid3Value = mid3.getText().trim();
        String tid3Value = tid3.getText().trim();
        if (!mid3Value.isEmpty() && !tid3Value.isEmpty()) {
            allValues.put(mid3Value, tid3Value);
        }
        String mid4Value = mid4.getText().trim();
        String tid4Value = tid4.getText().trim();
        if (!mid4Value.isEmpty() && !tid4Value.isEmpty()) {
            allValues.put(mid4Value, tid4Value);
        }
        String mid5Value = mid5.getText().trim();
        String tid5Value = tid5.getText().trim();
        if (!mid5Value.isEmpty() && !tid5Value.isEmpty()) {
            allValues.put(mid5Value, tid5Value);
        }

        return new ArrayList<>(allValues.entrySet());
    }

    private List<Map.Entry<String, String>> getVinattiTidList() {
        HashMap<String, String> tidList = new HashMap<>();

        String tid1Value = tid1.getText().trim();
        if (!tid1Value.isEmpty()) {
            tidList.put("1", tid1Value);
        }
        String tid2Value = tid2.getText().trim();
        if (!tid2Value.isEmpty()) {
            tidList.put("2", tid2Value);
        }
        String tid3Value = tid3.getText().trim();
        if (!tid3Value.isEmpty()) {
            tidList.put("3", tid3Value);
        }
        String tid4Value = tid4.getText().trim();
        if (!tid4Value.isEmpty()) {
            tidList.put("4", tid4Value);
        }
        String tid5Value = tid5.getText().trim();
        if (!tid5Value.isEmpty()) {
            tidList.put("5", tid5Value);
        }

        return new ArrayList<>(tidList.entrySet());
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
        replaces.put("{authorizerComId}", authorizerComId.getText().trim());
        replaces.put("{authorizerComIdDate}", authorizerComIdDate.getText().trim());
        replaces.put("{authorizerComIdPlace}", authorizerComIdPlace.getText().trim());
        replaces.put("{contractNo}", authorizerComId.getText().trim() + "/2024/HĐDV/ONEFIN");

        replaces.put("{authorizerName}", authorizerName.getText().trim().toUpperCase());
        replaces.put("{authorizerId}", authorizerId.getText().trim());
        replaces.put("{authorizerIdDate}", authorizerIdDate.getText().trim());
        replaces.put("{authorizerIdPlace}", authorizerIdPlace.getText().trim());

        String deviceName1Value = deviceName1.getSelectedItem().toString().trim();
        replaces.put("{deviceName1}", deviceName1Value.isEmpty() ? "" : (serviceType + deviceName1Value));
        String deviceName2Value = deviceName2.getSelectedItem().toString().trim();
        replaces.put("{deviceName2}", deviceName2Value.isEmpty() ? "" : (serviceType + deviceName2Value));
        replaces.put("{index2}", deviceName2Value.isEmpty() ? "" : "2");
        String deviceName3Value = deviceName3.getSelectedItem().toString().trim();
        replaces.put("{deviceName3}", deviceName3Value.isEmpty() ? "" : (serviceType + deviceName3Value));
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

        replaces.put("{handoverName}", handoverName.getSelectedItem().toString().trim().toUpperCase());
        replaces.put("{handoverId}", handoverId.getText().trim());
        replaces.put("{handoverIdDate}", handoverIdDate.getText().trim());
        replaces.put("{handoverIdPlace}", handoverIdPlace.getText().trim());

        StringBuilder mids = new StringBuilder();
        StringBuilder tids = new StringBuilder();
        StringBuilder midTitle = new StringBuilder();
        StringBuilder tidTitle = new StringBuilder();

        List<Map.Entry<String, String>> midTidList = getMidTidList();
        if (midTidList.size() > 1) {
            for (int i = 0; i < midTidList.size(); i++) {
                mids.append(midTidList.get(i).getKey()).append(" ".repeat(10));
                tids.append(midTidList.get(i).getValue()).append(" ".repeat(10));
                midTitle.append("MiD").append(" ").append(i + 1).append(":").append(" ");
                tidTitle.append("TiD").append(" ").append(i + 1).append(":").append(" ");
            }
        } else if (midTidList.size() == 1) {
            mids.append(midTidList.get(0).getKey());
            tids.append(midTidList.get(0).getValue());
            midTitle.append("MiD:");
            tidTitle.append("TiD:");
        }

        replaces.put("{mid}", midTitle.toString());
        replaces.put("{tid}", tidTitle.toString());
        replaces.put("{midList}", mids.toString());
        replaces.put("{tidList}", tids.toString());

        replaces.put("{tid1}", tid1.getText().trim());
        replaces.put("{authorizedBirthday}", authorizedBirthday.getText().trim());
        replaces.put("{vinattiNo}", vinattiNo.getText().trim());
        replaces.put("{authorizerTaxCode}", authorizerTaxCode.getText().trim());
        replaces.put("{authorizerTel}", authorizerTel.getText().trim());
        replaces.put("{authorizerEmail}", authorizerEmail.getText().trim());
        replaces.put("{shopName}", shopName.getText().trim().toUpperCase());

        replaces.put("{key0}", key0);
        replaces.put("{value0}", value0);
        replaces.put("{key1}", key1);
        replaces.put("{value1}", value1);
        replaces.put("{key2}", key2);
        replaces.put("{value2}", value2);
        replaces.put("{key3}", key3);
        replaces.put("{value3}", value3);
        replaces.put("{key4}", key4);
        replaces.put("{value4}", value4);

        return replaces;
    }

    private void replaceTextInDocxFile(String fileName, HashMap<String, String> replaces) {
        String srcDocx = System.getProperty("user.dir") + File.separator + "sourceDocs" + File.separator
                + fileName + ".docx";
        String dstDocx = System.getProperty("user.dir") + File.separator + "resultDocs" + File.separator
                + authorizedId.getSelectedItem().toString().trim() + "_" + fileName + new SimpleDateFormat("_yyyy-MM-dd").format(new Date()) + ".docx";
        replaceText(srcDocx, dstDocx, replaces);
    }

    private void exportAppotaExcel() {
        String srcAppotaExcel = System.getProperty("user.dir") + File.separator + "sourceDocs" + File.separator
                + "UY-QUYEN-APPOTA.xlsx";
        String dstAppotaExcel = System.getProperty("user.dir") + File.separator + "resultDocs" + File.separator
                + authorizedId.getSelectedItem().toString().trim() + "_UY-QUYEN-APPOTA" + new SimpleDateFormat("_yyyy-MM-dd").format(new Date()) + ".xlsx";

        List<String> dataList = new ArrayList<>();
        dataList.add("1");
        dataList.add(posName.getText().trim().toUpperCase());
        dataList.add(authorizerComAddress.getText().trim());
        dataList.add(mid1.getText().trim());
        dataList.add(tid1.getText().trim());
        dataList.add(serie.getText().trim());
        dataList.add("105881679913");
        dataList.add("NGUYỄN BÁ BA");
        dataList.add("Vietinbank");
        dataList.add(authorizedAcc.getText().trim());
        dataList.add(authorizedName.getSelectedItem().toString().trim().toUpperCase());
        dataList.add(authorizedBank.getText().trim());

        updateRecord(srcAppotaExcel, dstAppotaExcel, "Sheet1", 2, dataList);
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
            if (select == 1 || select == 3 || select == 4) {
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
        authorizedBirthday.setText("");
    }

    private List<String> saveCustomerInfo() {
        List<String> authorizedInfo = new ArrayList<>();
        authorizedInfo.add("index 0");
        authorizedInfo.add(authorizedName.getSelectedItem().toString().trim().toUpperCase());
        authorizedInfo.add(authorizedBirthday.getText().trim());
        authorizedInfo.add(authorizedTel.getText().trim());
        authorizedInfo.add(authorizedEmail.getText().trim());
        authorizedInfo.add(authorizedId.getSelectedItem().toString().trim());
        authorizedInfo.add(authorizedIdDate.getText().trim());
        authorizedInfo.add(authorizedIdPlace.getText().trim());
        authorizedInfo.add(authorizedAddress.getText().trim());
        authorizedInfo.add(authorizedAcc.getText().trim());
        authorizedInfo.add(authorizedBank.getText().trim());
        authorizedInfo.add("");
        return authorizedInfo;
    }

    private List<String> saveAuthorizerInfo() {
        List<String> authorizerInfo = new ArrayList<>();
        authorizerInfo.add("index 0");
        authorizerInfo.add(authorizerComName.getSelectedItem().toString().trim().toUpperCase());
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
        authorizerInfo.add(vinattiNo.getText().trim());
        return authorizerInfo;
    }

    private List<String> saveStaffInfo() {
        List<String> staffInfo = new ArrayList<>();
        staffInfo.add("index 0");
        staffInfo.add(handoverName.getSelectedItem().toString().trim().toUpperCase());
        staffInfo.add(handoverId.getText().trim());
        staffInfo.add(handoverIdDate.getText().trim());
        staffInfo.add(handoverIdPlace.getText().trim());
        return staffInfo;
    }

    private void addSuggestionListener(JComboBox<String> comboBox, ComboBoxFlags flags, String sheetName) {
        comboBox.getEditor().getEditorComponent().setForeground(new Color(0, 104, 139));

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
                if (flags.flag) {
                    return;
                }
                flags.flag = true;

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
                            comboBox.showPopup();
                        } else {
                            comboBox.hidePopup();
                        }
                    } else {
                        comboBox.hidePopup();
                    }

                    flags.flag = false;
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
                        flags.flag = true;
                        SwingUtilities.invokeLater(() -> flags.flag = false);
                        comboBox.dispatchEvent(e);
                    }
                }
            }
        });
    }
}

class ComboBoxFlags {
    boolean flag = false;
}