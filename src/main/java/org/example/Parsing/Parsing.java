package org.example.Parsing;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

import java.io.*;

public class Parsing {
    public static void parse() {
        String url = "https://cbr.ru/currency_base/daily/";
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("course");

            Document document = Jsoup.connect(url).get();
            Elements nameOfSheet = document.select("div.table-wrapper > div.table > table.data > tbody > tr >" +
                    " th");
            Elements value = document.select("div.table-wrapper > div.table > table.data > tbody > " +
                    "tr :not(th)");

            int numsRow = 0;

            Row header = sheet.createRow(numsRow);
            header.createCell(0).setCellValue(nameOfSheet.get(0).text());
            header.createCell(1).setCellValue(nameOfSheet.get(1).text());
            header.createCell(2).setCellValue(nameOfSheet.get(2).text());
            header.createCell(3).setCellValue(nameOfSheet.get(3).text());
            header.createCell(4).setCellValue(nameOfSheet.get(4).text());
            numsRow++;

            sheet.autoSizeColumn(0);
            sheet.setColumnWidth(3, 9000);

            int numsIndex = 0;

            Row dataRow1 = sheet.createRow(numsRow);
            dataRow1.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow1.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow1.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow1.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow1.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow2 = sheet.createRow(numsRow);
            dataRow2.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow2.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow2.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow2.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow2.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow3 = sheet.createRow(numsRow);
            dataRow3.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow3.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow3.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow3.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow3.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow4 = sheet.createRow(numsRow);
            dataRow4.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow4.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow4.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow4.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow4.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow5 = sheet.createRow(numsRow);
            dataRow5.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow5.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow5.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow5.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow5.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow6 = sheet.createRow(numsRow);
            dataRow6.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow6.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow6.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow6.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow6.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow7 = sheet.createRow(numsRow);
            dataRow7.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow7.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow7.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow7.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow7.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow8 = sheet.createRow(numsRow);
            dataRow8.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow8.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow8.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow8.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow8.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow9 = sheet.createRow(numsRow);
            dataRow9.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow9.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow9.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow9.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow9.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow10 = sheet.createRow(numsRow);
            dataRow10.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow10.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow10.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow10.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow10.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow11 = sheet.createRow(numsRow);
            dataRow11.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow11.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow11.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow11.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow11.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow12 = sheet.createRow(numsRow);
            dataRow12.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow12.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow12.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow12.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow12.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow13 = sheet.createRow(numsRow);
            dataRow13.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow13.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow13.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow13.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow13.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow14 = sheet.createRow(numsRow);
            dataRow14.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow14.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow14.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow14.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow14.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow15 = sheet.createRow(numsRow);
            dataRow15.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow15.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow15.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow15.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow15.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow16 = sheet.createRow(numsRow);
            dataRow16.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow16.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow16.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow16.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow16.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow17 = sheet.createRow(numsRow);
            dataRow17.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow17.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow17.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow17.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow17.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow18 = sheet.createRow(numsRow);
            dataRow18.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow18.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow18.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow18.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow18.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow19 = sheet.createRow(numsRow);
            dataRow19.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow19.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow19.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow19.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow19.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow20 = sheet.createRow(numsRow);
            dataRow20.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow20.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow20.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow20.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow20.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow21 = sheet.createRow(numsRow);
            dataRow21.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow21.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow21.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow21.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow21.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow22 = sheet.createRow(numsRow);
            dataRow22.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow22.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow22.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow22.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow22.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow23 = sheet.createRow(numsRow);
            dataRow23.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow23.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow23.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow23.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow23.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow24 = sheet.createRow(numsRow);
            dataRow24.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow24.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow24.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow24.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow24.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow25 = sheet.createRow(numsRow);
            dataRow25.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow25.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow25.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow25.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow25.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow26 = sheet.createRow(numsRow);
            dataRow26.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow26.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow26.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow26.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow26.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow27 = sheet.createRow(numsRow);
            dataRow27.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow27.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow27.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow27.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow27.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow28 = sheet.createRow(numsRow);
            dataRow28.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow28.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow28.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow28.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow28.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow29 = sheet.createRow(numsRow);
            dataRow29.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow29.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow29.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow29.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow29.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow30 = sheet.createRow(numsRow);
            dataRow30.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow30.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow30.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow30.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow30.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow31 = sheet.createRow(numsRow);
            dataRow31.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow31.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow31.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow31.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow31.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow32 = sheet.createRow(numsRow);
            dataRow32.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow32.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow32.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow32.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow32.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow33 = sheet.createRow(numsRow);
            dataRow33.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow33.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow33.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow33.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow33.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow34 = sheet.createRow(numsRow);
            dataRow34.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow34.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow34.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow34.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow34.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow35 = sheet.createRow(numsRow);
            dataRow35.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow35.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow35.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow35.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow35.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow36 = sheet.createRow(numsRow);
            dataRow36.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow36.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow36.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow36.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow36.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow37 = sheet.createRow(numsRow);
            dataRow37.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow37.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow37.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow37.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow37.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow38 = sheet.createRow(numsRow);
            dataRow38.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow38.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow38.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow38.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow38.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow39 = sheet.createRow(numsRow);
            dataRow39.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow39.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow39.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow39.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow39.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow40 = sheet.createRow(numsRow);
            dataRow40.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow40.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow40.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow40.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow40.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow41 = sheet.createRow(numsRow);
            dataRow41.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow41.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow41.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow41.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow41.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow42 = sheet.createRow(numsRow);
            dataRow42.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow42.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow42.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow42.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow42.createCell(4).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            numsRow++;

            Row dataRow43 = sheet.createRow(numsRow);
            dataRow43.createCell(0).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow43.createCell(1).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow43.createCell(2).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow43.createCell(3).setCellValue(value.get(numsIndex).text());
            numsIndex++;
            dataRow43.createCell(4).setCellValue(value.get(numsIndex).text());

            String filePath = "excel.xlsx";
            FileOutputStream fileOutputStream = new FileOutputStream(filePath);
            workbook.write(fileOutputStream);
            workbook.close();

        } catch (Exception e) {
            System.out.println("Ошибка " + e);
        }
    }

}