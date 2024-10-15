package actions;

import java.text.NumberFormat;
import java.util.Locale;

public class NumberToWordsVN {

    private static final String[] units = {"không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"};
    private static final String[] tens = {"", "mười", "hai mươi", "ba mươi", "bốn mươi", "năm mươi", "sáu mươi", "bảy mươi", "tám mươi", "chín mươi"};
    private static final String[] scales = {"", "nghìn", "triệu", "tỷ"};

    public static String convertToWords(int number) {
        if (number == 0) {
            return "Không";
        }

        String numberStr = NumberFormat.getInstance(Locale.US).format(number);
        String[] parts = numberStr.split(",");

        StringBuilder words = new StringBuilder();
        int scaleIndex = parts.length - 1;

        for (String part : parts) {
            int partInt = Integer.parseInt(part);
            if (partInt > 0) {
                words.append(convertThreeDigitsToWords(partInt)).append(" ").append(scales[scaleIndex]).append(" ");
            }
            scaleIndex--;
        }

        String result = words.toString().trim();
        return capitalizeFirstLetter(result);
    }

    private static String convertThreeDigitsToWords(int number) {
        int hundreds = number / 100;
        int tensUnits = number % 100;
        int tens = tensUnits / 10;
        int units = tensUnits % 10;

        StringBuilder words = new StringBuilder();

        if (hundreds != 0) {
            words.append(NumberToWordsVN.units[hundreds]).append(" trăm ");
        }

        if (tensUnits != 0) {
            if (tensUnits < 10) {
                if (units == 5) {
                    words.append("lẻ năm");
                } else {
                    words.append("lẻ ").append(NumberToWordsVN.units[units]);
                }
            } else {
                words.append(NumberToWordsVN.tens[tens]).append(" ");
                if (units != 0) {
                    if (units == 5) {
                        words.append("lăm");
                    } else {
                        words.append(NumberToWordsVN.units[units]);
                    }
                }
            }
        }

        return words.toString().trim();
    }

    private static String capitalizeFirstLetter(String input) {
        if (input == null || input.isEmpty()) {
            return input;
        }
        return input.substring(0, 1).toUpperCase() + input.substring(1);
    }

    public static void main(String[] args) {
        int number = 120005000;
        System.out.println(convertToWords(number));
    }
}

