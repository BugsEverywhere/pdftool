package indi.simon.pdftool;

import com.spire.pdf.PdfDocument;
import com.spire.pdf.PdfPageBase;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;
import java.util.HashMap;
import java.util.Map;

public class Main {

    public static void main(String[] args) {

        File currentDir = new File("./");
        File[] allFils = currentDir.listFiles();
        if (allFils == null || allFils.length == 0) {
            throw new IllegalArgumentException("empty directory! what exactly do you want me to do with this?");
        }
        File excelFile = null;
        Map<String, Map<String, Map<String, Float>>> rateMap = new HashMap<>();
        for (File singleFile : allFils) {
            String fileName = singleFile.getName();
            System.out.println("file name:" + fileName);
            if (fileName.endsWith("pdf")) {
                String[] nameArr = fileName.split(" ");
                String nameDateWithSuffix = nameArr[nameArr.length - 1];
                String[] nameDateArr = nameDateWithSuffix.split("\\.");
                String dateStr = nameDateArr[0] + nameDateArr[1] + nameDateArr[2];
                Map<String, Map<String, Float>> singleDate = new HashMap<>();
                rateMap.put(dateStr, singleDate);
                PdfDocument doc = new PdfDocument();
                doc.loadFromFile(singleFile.getAbsolutePath());
                StringBuilder sb = new StringBuilder();
                PdfPageBase page;

                page = doc.getPages().get(0);
                sb.append(page.extractText(true));

                String content = sb.toString();
                String[] lines = content.split("\n");
                boolean startOfTable2 = false;
                int count = 11;
                for (String singleLine : lines) {
                    if (singleLine.equals("\r")) {
                        continue;
                    }
                    if (singleLine.contains("ONSHORE FCY INTERBANK BID")) {
                        startOfTable2 = true;
                        continue;
                    }
                    if (startOfTable2) {
                        if (count == 11) {
                            count--;
                            continue;
                        }
                        if (count > 0) {
                            singleLine = singleLine.trim();
                            //System.out.println("singleLine of PDF table:" + singleLine);
                            String[] cells = singleLine.split("\\s+");
                            if (cells.length < 8) {
                                continue;
                            }
                            for (int i = 0; i < cells.length; i++) {
                                cells[i] = cells[i].trim();
                            }

                            Map<String, Float> tenorMap = new HashMap<>();
                            singleDate.put(cells[0], tenorMap);
                            tenorMap.put("1W", Float.parseFloat(cells[2].substring(0, 4)));
                            tenorMap.put("2W", Float.parseFloat(cells[3].substring(0, 4)));
                            tenorMap.put("1M", Float.parseFloat(cells[4].substring(0, 4)));
                            tenorMap.put("3M", Float.parseFloat(cells[5].substring(0, 4)));
                            tenorMap.put("6M", Float.parseFloat(cells[6].substring(0, 4)));
                            tenorMap.put("12M", Float.parseFloat(cells[7].substring(0, 4)));
                            count--;
                        }
                    }
                }
                doc.close();
            } else if (fileName.endsWith("xls") || fileName.endsWith("xlsx")) {
                excelFile = singleFile;
            }
        }
        OutputStream os = null;
        try {
            if (excelFile == null) {
                return;
            }
            Workbook wb = getExcel(excelFile.getAbsolutePath());
            Sheet sheet = wb.getSheet("Sheet1");
            int rowNum = sheet.getLastRowNum();
            System.out.println("total row count:" + rowNum);
            for (int i = 1; i <= rowNum; i++) {
                Row row = sheet.getRow(i);
                Cell cell = row.getCell(5);
                String currency = null;
                try {
                    currency = cell.toString().trim();
                } catch (Exception e) {
                    System.out.println("currency error line num:" + i);
                    continue;
                }
                //System.out.println("currency:" + currency);
                if (currency.equals("") || currency.length() == 0) {
                    break;
                }
                //DecimalFormat df = new DecimalFormat("0");
                String date = row.getCell(13).toString();
                int tenor = (int) Double.parseDouble(row.getCell(15).toString());
                String tenorDay = row.getCell(16).toString();
                //System.out.println("currency:" + currency + ", value date:" + date + ", tenor:" + tenor + ", tenorDay:" + tenorDay);
                Map<String, Map<String, Float>> dateCurrencyMap = rateMap.get(date);
                if (dateCurrencyMap == null) {
                    System.out.println("lack of data in date:" + date + ", skip");
                    continue;
                }
                Map<String, Float> tenorRateMap = dateCurrencyMap.get(currency);
                String rate = null;
                DecimalFormat rateDf = new DecimalFormat("0.00");
                switch (tenor) {
                    case 12:
                        rate = rateDf.format(tenorRateMap.get("12M"));
                        break;
                    case 13:
                        rate = rateDf.format(tenorRateMap.get("12M"));
                        break;
                    case 6:
                        rate = rateDf.format(tenorRateMap.get("6M"));
                        break;
                    case 3:
                        rate = rateDf.format(tenorRateMap.get("3M"));
                        break;
                    case 14:
                        if (Double.parseDouble(tenorDay) == 14 || Double.parseDouble(tenorDay) == 15) {
                            rate = rateDf.format(tenorRateMap.get("2W"));
                        }
                        break;
                    case 1:
                        rate = rateDf.format(tenorRateMap.get("1M"));
                        break;
                    default:
                        rate = "0.0";
                }
                Cell cell1 = row.createCell(10);
                cell1.setCellValue(rate);
            }
            os = new FileOutputStream(excelFile);
            wb.write(os);
            wb.close();
        } catch (IOException e) {
            System.out.println(e.getMessage());
        } finally {
            if (os != null) {
                try {
                    os.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

    }
//        Gson gson = new GsonBuilder().setPrettyPrinting().create();
//        System.out.println(gson.toJson(rateMap));


    public static Workbook getExcel(String filePath) {
        Workbook wb = null;
        File file = new File(filePath);
        if (!file.exists()) {
            System.out.println("file not exist");
            wb = null;
        } else {
            String fileType = filePath.substring(filePath.lastIndexOf("."));
            InputStream is = null;
            try {
                is = new FileInputStream(filePath);
                if (".xls".equals(fileType)) {
                    wb = new HSSFWorkbook(is);
                } else if (".xlsx".equals(fileType)) {
                    ZipSecureFile.setMinInflateRatio(-1.0d);
                    wb = new XSSFWorkbook(is);
                } else {
                    System.out.println("incorrect format");
                    wb = null;
                }
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                if (is != null) {
                    try {
                        is.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }
        }
        return wb;
    }

}
