package indi.simon.pdftool;

public class AnotherMain {

    public static void main(String[] args) {





    }

    public int getLength(String str) {
        if (str == null || str.length() < 1) {
            return 0;
        }
        boolean[][] result = new boolean[str.length()][str.length()];
        for (int i = 0; i < str.length(); i++) {
            result[i][i] = true;
        }
        char[] chars = str.toCharArray();
        int maxLength = 1;
        for (int i = 1; i < chars.length; i++) {
            for (int j = 0; j < i; j++) {
                result[i][j] = chars[i] == chars[j] && (i - j == 1 || result[i - 1][j + 1]);
                if (result[i][j]) {
                    maxLength = Math.max(maxLength, i - j + 1);
                }
            }
        }
        System.out.println(maxLength);
        return maxLength;
    }
}
