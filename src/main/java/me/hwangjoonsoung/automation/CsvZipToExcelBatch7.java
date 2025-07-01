package me.hwangjoonsoung.automation;

import com.opencsv.CSVParser;
import com.opencsv.CSVParserBuilder;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import com.opencsv.exceptions.CsvException;
import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipArchiveInputStream;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.mozilla.universalchardet.UniversalDetector;

import java.io.*;
import java.nio.charset.Charset;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

// 일자별 요일별 파워링크까지 한번에 동작
public class CsvZipToExcelBatch7 {

    public static void main(String[] args) throws Exception {
        File zipFile = new File("src/main/java/me/hwangjoonsoung/automation/inputCSVZip/archives.zip");
        File unzipDir = new File("build/unzipped_place");
        File templateFile = new File("src/main/java/me/hwangjoonsoung/automation/basedExcelFile/12월 키워드보고서.xlsx");
        File outputDir = new File("build/output_place");

        if (!outputDir.exists()) outputDir.mkdirs();

        extractZip(zipFile, unzipDir);
        processAllCsvSet(unzipDir, templateFile.getAbsolutePath(), outputDir);
    }

    public static void extractZip(File zipFile, File destDir) throws IOException {
        try (ZipArchiveInputStream zis = new ZipArchiveInputStream(
                new FileInputStream(zipFile), "UTF-8", false, true)) {
            ZipArchiveEntry entry;
            while ((entry = zis.getNextZipEntry()) != null) {
                File outFile = new File(destDir, entry.getName());
                outFile.getParentFile().mkdirs();
                try (FileOutputStream fos = new FileOutputStream(outFile)) {
                    IOUtils.copy(zis, fos);
                }
            }
        }
        System.out.println("✅ 압축 해제 완료 (EUC-KR 해석): " + destDir.getAbsolutePath());
    }

    public static void processAllCsvSet(File folder, String templatePath, File outputDir) throws Exception {
        File[] files = folder.listFiles((dir, name) -> name.endsWith(".csv"));
        if (files == null) return;

        Set<String> idSet = new HashSet<>();
        for (File f : files) {
            String name = f.getName();
            if (name.startsWith("파워링크보고서,")) {
                String id = name.replace("파워링크보고서,", "").replace(".csv", "");
                idSet.add(id);
            }
        }

        for (String id : idSet) {
            File daily = new File(folder, "일별보고서," + id + ".csv");
            File time = new File(folder, "요일별보고서," + id + ".csv");
            File outputFile = new File(outputDir, "12월_키워드보고서_" + id + ".xlsx");

            if (daily.exists()) {
                processOneSet(daily, time, templatePath, outputFile);
            } else {
                System.out.println("❌ 일별 파일 누락: " + id);
            }
        }
    }

    public static void processOneSet(File dailyCsv, File timeCsv, String templatePath, File outputFile) throws Exception {
        String baseName = outputFile.getName().replace("12월_키워드보고서_", "").replace(".xlsx", "");
        File powerlinkCsv = new File("build/unzipped_place/파워링크보고서," + baseName + ".csv");
        File shoppingCsv = new File("build/unzipped_place/쇼핑검색보고서," + baseName + ".csv");
        File placeCsv = new File("build/unzipped_place/플레이스보고서," + baseName + ".csv");
        FileInputStream fis = new FileInputStream(templatePath);
        Workbook workbook = new XSSFWorkbook(fis);

        Sheet dailySheet = workbook.getSheet("일자별");
        Sheet shoppingSheet = workbook.getSheet("쇼핑검색");

        writeDailySheet(dailySheet, dailyCsv, workbook);

       if (timeCsv.exists()) {
            Sheet timeSheet = workbook.getSheet("시간별");
            writeTimeSheet(timeSheet, timeCsv, workbook);
        }

         //done
        if (powerlinkCsv.exists()) {
            Sheet powerlinkSheet = workbook.getSheet("파워링크");
            writePowerlinkSheet(powerlinkSheet, powerlinkCsv, workbook);
        }

        if (shoppingCsv.exists()) {
            writeShoppingSheet(shoppingSheet, shoppingCsv, workbook);
        }

        if (placeCsv.exists()) {
            Sheet placeSheet = workbook.getSheet("플레이스");
            writePlaceSheet(placeSheet, placeCsv, workbook);
        }

        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
            workbook.write(fos);
        }

        workbook.close();
        System.out.println("✅ 저장 완료: " + outputFile.getAbsolutePath());
    }

    public static void writeDailySheet(Sheet sheet, File csvFile, Workbook wb) throws IOException, CsvException {
        String encoding = detectEncoding(csvFile);
        try (CSVReader reader = new CSVReader(
                new InputStreamReader(new FileInputStream(csvFile), Charset.forName(encoding)))) {

            List<String[]> rows = reader.readAll();
            int startRow = 28;
            int startCol = 40;

            for (int i = 2; i < rows.size(); i++) {
                String[] row = rows.get(i);
                Row excelRow = sheet.getRow(startRow);
                if (excelRow == null) excelRow = sheet.createRow(startRow);

                for (int j = 0; j < 9; j++) {
                    Cell cell = excelRow.createCell(startCol + j);
                    String val = row[j].trim();

                    try {
                        // 숫자 입력 시 숫자로 넣되, 스타일은 지정하지 않음
                        cell.setCellValue(Double.parseDouble(val.replace(",", "")));
                    } catch (NumberFormatException e) {
                        cell.setCellValue(val);  // 문자열 그대로
                    }
                    // cell.setCellStyle(...) 생략 → "일반" 유지
                }
                startRow++;
            }
        }
    }

    public static void writeTimeSheet(Sheet sheet, File csvFile, Workbook wb) throws IOException, CsvException {
        String encoding = detectEncoding(csvFile);
        try (CSVReader reader = new CSVReader(
                new InputStreamReader(new FileInputStream(csvFile), Charset.forName(encoding)))) {

            List<String[]> rows = reader.readAll();
            int startRow = 60;  // Excel 기준 61행
            int startCol = 1;   // Excel 기준 B열

            DataFormat format = wb.createDataFormat();

            // 일반 서식
            CellStyle generalStyle = wb.createCellStyle();
            generalStyle.setDataFormat(format.getFormat("General"));

            // 회계 서식
            CellStyle accountingStyle = wb.createCellStyle();
            accountingStyle.setDataFormat((short) 44);

            for (int i = 2; i < rows.size(); i++) {
                String[] row = rows.get(i);
                Row excelRow = sheet.getRow(startRow);
                if (excelRow == null) {
                    excelRow = sheet.createRow(startRow);
                }

                for (int j = 0; j < 9; j++) {
                    Cell cell = excelRow.createCell(startCol + j);
                    String val = row[j].trim();
                    try {
                        cell.setCellValue(Double.parseDouble(val.replace(",", "")));
                    } catch (NumberFormatException e) {
                        cell.setCellValue(val);  // 원본 텍스트 그대로
                    }
                }
                startRow++;
            }
        }
    }

    public static void writePowerlinkSheet(Sheet sheet, File csvFile, Workbook wb) throws IOException, CsvException {
        String encoding = detectEncoding(csvFile);
        try (CSVReader reader = new CSVReader(new InputStreamReader(new FileInputStream(csvFile), Charset.forName(encoding)))) {
            List<String[]> rows = reader.readAll();
            int startRow = 28;  // Excel 29행
            int startCol = 1;   // Excel B열

            DataFormat format = wb.createDataFormat();

            // 일반 서식
            CellStyle generalStyle = wb.createCellStyle();
            generalStyle.setDataFormat(format.getFormat("General"));

            // 회계 서식
            CellStyle accountingStyle = wb.createCellStyle();
            accountingStyle.setDataFormat((short) 44);

            for (int i = 2; i < rows.size(); i++) {  // 6행부터 시작
                String[] row = rows.get(i);
                Row excelRow = sheet.getRow(startRow);
                if (excelRow == null) {
                    excelRow = sheet.createRow(startRow);
                }

                for (int j = 3; j <= 13; j++) {
                    Cell cell = excelRow.createCell(startCol + (j -3));
                    String val = row[j].trim();
                    try {
                        cell.setCellValue(Double.parseDouble(val.replace(",", "")));
                    } catch (NumberFormatException e) {
                        cell.setCellValue(val);  // 원본 텍스트 그대로
                    }
                }
                startRow++;
            }
        }
    }

    public static void writeShoppingSheet(Sheet sheet, File csvFile, Workbook wb) throws IOException, CsvException {
        String encoding = detectEncoding(csvFile);

        CSVParser parser = new CSVParserBuilder()
                .withSeparator(',')         // CSV 구분자: 쉼표
                .withQuoteChar('"')         // 인용문자: "
                .withEscapeChar(CSVParser.NULL_CHARACTER) // ✅ 이스케이프 문자 제거
                .build();

        try (CSVReader reader = new CSVReaderBuilder(new InputStreamReader(new FileInputStream(csvFile), Charset.forName(encoding)))
                .withCSVParser(parser)
                .build()) {

            List<String[]> rows = reader.readAll();
            int startRow = 28; // Excel 기준 29행
            int startCol = 1;  // Excel B열

            for (int i = 2; i < rows.size(); i++) {
                String[] row = rows.get(i);
                if (row.length < 12) {
                    System.out.printf("⚠️"+csvFile.getName()+"파일 ⚠️ Skipping row at index %d: too short (length = %d)%n", i, row.length);
                    continue;
                }
                if (!"쇼핑검색".equals(row[0])) continue;

                Row excelRow = sheet.getRow(startRow);
                if (excelRow == null) excelRow = sheet.createRow(startRow);

                for (int j = 1; j <= 11; j++) {
                    Cell cell = excelRow.createCell(startCol + (j-1));
                    String val = row[j].replace(",", "").trim();
                    try {
                        cell.setCellValue(Double.parseDouble(val.replace(",", "")));
                    } catch (NumberFormatException e) {
                        System.out.println(val+" : "+ e);
                        cell.setCellValue(val);  // 원본 텍스트 그대로
                    }

                }
                startRow++;
            }
        }
    }

    public static void writePlaceSheet(Sheet sheet, File csvFile, Workbook wb) throws IOException, CsvException {
//        System.out.println("플레이스 보고서 작성중---"+csvFile.getName()+"---파일");
        String encoding = detectEncoding(csvFile);

        CSVParser parser = new CSVParserBuilder()
                .withSeparator(',')         // CSV 구분자: 쉼표
                .withQuoteChar('"')         // 인용문자: "
                .withEscapeChar(CSVParser.NULL_CHARACTER) // ✅ 이스케이프 문자 제거
                .build();

        try (CSVReader reader = new CSVReaderBuilder(
                new InputStreamReader(new FileInputStream(csvFile), Charset.forName(encoding)))
                .withCSVParser(parser)
                .build()) {

            List<String[]> rows = reader.readAll();
            int startRow = 28; // Excel 기준 29행
            int startCol = 1;  // Excel C열 (index 1)

            DataFormat format = wb.createDataFormat();

            // 일반 서식
            CellStyle generalStyle = wb.createCellStyle();
            generalStyle.setDataFormat(format.getFormat("General"));

            // 회계 서식
            CellStyle accountingStyle = wb.createCellStyle();
            accountingStyle.setDataFormat((short) 44);

            for (int i = 2; i < rows.size(); i++) {
                String[] row = rows.get(i);

                if (row.length < 10) {
                    System.out.printf("⚠️"+csvFile.getName()+"파일 ⚠️ Skipping row at index %d: too short (length = %d)%n", i, row.length);
                    continue;
                }

                String campaign = row[0].replaceAll("\"", "").trim();
                if (!"플레이스".equals(campaign)) continue;

                Row excelRow = sheet.getRow(startRow);
                if (excelRow == null) excelRow = sheet.createRow(startRow);

                for (int j = 0; j < 10; j++) {
                    Cell cell = excelRow.createCell(startCol + j);
                    String val = row[j].replace(",", "").trim();

                    try {
                        double num = Double.parseDouble(val);
                        cell.setCellValue(num);

                    } catch (NumberFormatException e) {
                        cell.setCellValue(val);
                    }
                }

                startRow++;
            }
        }
    }

    public static String detectEncoding(File file) throws IOException {
        byte[] buf = new byte[4096];
        FileInputStream fis = new FileInputStream(file);
        UniversalDetector detector = new UniversalDetector(null);

        int nread;
        while ((nread = fis.read(buf)) > 0 && !detector.isDone()) {
            detector.handleData(buf, 0, nread);
        }
        detector.dataEnd();
        fis.close();

        String encoding = detector.getDetectedCharset();
        return encoding != null ? encoding : "EUC-KR";
    }

    // 🔁 글자 색상을 흰색으로 바꾸고 싶다면 아래처럼 변경하세요:
    // greenFont.setColor(IndexedColors.WHITE.getIndex());
}
