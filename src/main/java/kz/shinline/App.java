package kz.shinline;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import javax.swing.*;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Locale;

public class App {

    private JPanel thePanel;

    public static void main(String[] args) {
        JFrame frame = new JFrame("Загрузка");
        frame.setContentPane(new App().thePanel);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);

        SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy", Locale.ENGLISH);
        String[] currencies = {"20", "16", "10"};
        final String today = dateFormat.format(new Date());
        ArrayList<Boolean> progress = new ArrayList<>();
        for (final String currency : currencies) {
            try {
                String quantityText = "";
                String currencyName = "";
                double currencyAmount = 0.0;
                String fileName = "";
                switch (currency) {
                    case "20": {
                        quantityText = "UZS_quant";
                        currencyName = "UZS";
                        fileName = "UZS rates.xlsx";
                        currencyAmount = 100.0;
                        break;
                    }
                    case "16": {
                        quantityText = "RUB_quant";
                        currencyName = "RUB";
                        fileName = "RUB rates.xlsx";
                        currencyAmount = 1.0;
                        break;
                    }
                    case "10": {
                        quantityText = "KGS_quant";
                        currencyName = "KGS";
                        fileName = "KGS rates.xlsx";
                        currencyAmount = 1.0;
                        break;
                    }
                }
                Workbook workbook = new XSSFWorkbook();
                Sheet sheet = workbook.createSheet("Courses");
                Row firstRow = sheet.createRow(0);
                firstRow.createCell(0).setCellValue("Date");
                firstRow.createCell(1).setCellValue(quantityText);
                firstRow.createCell(2).setCellValue(currencyName);
                Document doc = Jsoup.connect("https://www.nationalbank.kz/ru/exchangerates/ezhednevnye-oficialnye-rynochnye-kursy-valyut/report?rates%5B%5D="+
                        currency +"&beginDate=01.01.2016&endDate=" + today).get();
                Element table = doc.select("table").get(0);
                Elements rows = table.select("tr");
                int rowIndex = 1;
                for (Element row : rows) {
                    Elements column = row.select("td");
                    if (column.size() > 0) {
                        Row aRow = sheet.createRow(rowIndex++);
                        aRow.createCell(0).setCellValue(column.get(0).text());
                        aRow.createCell(1).setCellValue(currencyAmount);
                        aRow.createCell(2).setCellValue(Double.parseDouble(column.get(2).text()));
                    }
                }
                FileOutputStream fileOut = new FileOutputStream(fileName);
                workbook.write(fileOut);
                fileOut.close();
                workbook.close();
                progress.add(true);
            } catch (Exception e) {
                e.printStackTrace();
                progress.add(false);
            }
            if (progress.size() == 3) {
                System.exit(0);
            }
        }
    }
}



