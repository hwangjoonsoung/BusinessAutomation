package me.hwangjoonsoung.automation;

import com.opencsv.CSVParser;
import com.opencsv.CSVParserBuilder;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import com.opencsv.exceptions.CsvException;
import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipArchiveInputStream;
import org.apache.commons.compress.harmony.pack200.NewAttribute;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.mozilla.universalchardet.UniversalDetector;

import java.io.*;
import java.nio.charset.Charset;
import java.util.*;

// todo: xlxm -> xlsx 작업 필요
// todo: 경로 변경 작업 필요
// 일자별 요일별 파워링크까지 한번에 동작
public class CsvZipToExcelBatch7 {

    static LinkedHashSet linkedHashSet = new LinkedHashSet();

    public static void main(String[] args) throws Exception {
        File zipFile = new File("src/main/java/me/hwangjoonsoung/automation/inputCSVZip/archives.zip");
        File unzipDir = new File("build/unzipped_place");
        File templateFile = new File("src/main/java/me/hwangjoonsoung/automation/basedExcelFile/06월 키워드보고서.xlsm");
        File outputDir = new File("build/output_place");

        if (!outputDir.exists()) outputDir.mkdirs();

        extractZip(zipFile, unzipDir);
        processAllCsvSet(unzipDir, templateFile.getAbsolutePath(), outputDir);

        System.out.println("===========================================\n뭔가 이상한 파일들 : " + linkedHashSet);
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
            if (name.startsWith("일별보고서,")) {
                String id = name.replace("일별보고서,", "").replace(".csv", "");
                idSet.add(id);
            }
        }

        for (String id : idSet) {
            File daily = new File(folder, "일별보고서," + id + ".csv");
            File time = new File(folder, "요일별보고서," + id + ".csv");
            File outputFile = new File(outputDir, "06월_키워드보고서_" + id + ".xlsm");

//            if(!id.equals("bogangwood_naver")){
//                continue;
//            }
            if (daily.exists()) {
                processOneSet(daily, time, templatePath, outputFile);
            } else {
                System.out.println("❌ 일별 파일 누락: " + id);
            }
        }
    }

    public static void processOneSet(File dailyCsv, File timeCsv, String templatePath, File outputFile) throws Exception {
        String baseName = outputFile.getName().replace("06월_키워드보고서_", "").replace(".xlsm", "");
        File powerlinkCsv = new File("build/unzipped_place/파워링크보고서," + baseName + ".csv");
        File shoppingCsv = new File("build/unzipped_place/쇼핑검색보고서," + baseName + ".csv");
        File placeCsv = new File("build/unzipped_place/플레이스보고서," + baseName + ".csv");
        FileInputStream fis = new FileInputStream(templatePath);
        Workbook workbook = new XSSFWorkbook(fis);

        Sheet dailySheet = workbook.getSheet("일자별");
        writeDailySheet(dailySheet, dailyCsv, workbook);

        if (timeCsv.exists()) {
            Sheet timeSheet = workbook.getSheet("시간별");
            writeTimeSheet(timeSheet, timeCsv, workbook);
        } else {
            Sheet timeSheet = workbook.getSheet("시간별");
            if (timeSheet != null) workbook.removeSheetAt(workbook.getSheetIndex(timeSheet));
        }

        if (powerlinkCsv.exists()) {
            Sheet powerlinkSheet = workbook.getSheet("파워링크");
            writePowerlinkSheet(powerlinkSheet, powerlinkCsv, workbook);
        } else {
            Sheet powerlinkSheet = workbook.getSheet("파워링크");
            if (powerlinkSheet != null) workbook.removeSheetAt(workbook.getSheetIndex(powerlinkSheet));
        }

        if (shoppingCsv.exists()) {
            Sheet shoppingSheet = workbook.getSheet("쇼핑검색");
            writeShoppingSheet(shoppingSheet, shoppingCsv, workbook);
        } else {
            Sheet shoppingSheet = workbook.getSheet("쇼핑검색");
            if (shoppingSheet != null) workbook.removeSheetAt(workbook.getSheetIndex(shoppingSheet));
        }

        if (placeCsv.exists()) {
            Sheet placeSheet = workbook.getSheet("플레이스");
            writePlaceSheet(placeSheet, placeCsv, workbook);
        } else {
            Sheet placeSheet = workbook.getSheet("플레이스");
            if (placeSheet != null) workbook.removeSheetAt(workbook.getSheetIndex(placeSheet));
        }
        writeCoverSheet(workbook, baseName);

        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
            workbook.write(fos);
        }

        workbook.close();
        System.out.println("✅ 작업완료 완료: " + outputFile.getAbsolutePath());
    }

    public static void writeCoverSheet(Workbook wb, String baseName) {
        Sheet coverSheet = wb.getSheet("표지");
        if (coverSheet == null) {
            System.out.println("⚠️ 표지 시트가 존재하지 않습니다.");
            return;
        }

        // C4 셀 위치는 (row 3, column 2) → 0-based index
        Row row = coverSheet.getRow(3);
        if (row == null) {
            row = coverSheet.createRow(3);
        }

        Cell cell = row.getCell(2);
        if (cell == null) {
            cell = row.createCell(2);
        }

        cell.setCellValue(baseName);
    }

    public static void writeDailySheet(Sheet sheet, File csvFile, Workbook wb) throws IOException, CsvException {
        System.out.println("일자별 작업중 csv file name = " + csvFile);
        String encoding = detectEncoding(csvFile);
        try (CSVReader reader = new CSVReader(
                new InputStreamReader(new FileInputStream(csvFile), Charset.forName(encoding)))) {

            List<String[]> rows = reader.readAll();
            int startRow = 28;
            int startCol = 40;
            String csvValue = "";
            boolean isSomethingWrongFile = false;

            for (int i = 2; i < rows.size(); i++) {
                String[] row = rows.get(i);
                Row excelRow = sheet.getRow(startRow);
                if (excelRow == null) excelRow = sheet.createRow(startRow);

                for (int j = 0; j < 9; j++) {
                    Cell cell = excelRow.createCell(startCol + j);

                    if (row.length < 9 && j >= row.length) {
                        if(!isSomethingWrongFile){
                            linkedHashSet.add(csvFile.getName());
                        }
                        csvValue = "0";
                    } else {
                        csvValue = row[j];
                    }
                    String val = csvValue.trim();

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
        System.out.println("시간별 작업중 csv file name = " + csvFile);
        String encoding = detectEncoding(csvFile);
        try (CSVReader reader = new CSVReader(
                new InputStreamReader(new FileInputStream(csvFile), Charset.forName(encoding)))) {

            List<String[]> rows = reader.readAll();
            int startRow = 60;  // Excel 기준 61행
            int startCol = 1;   // Excel 기준 B열

            DataFormat format = wb.createDataFormat();

            // 일반 서식
            CellStyle generalStyle = wb.createCellStyle();
            generalStyle.setBorderTop(BorderStyle.THIN);
            generalStyle.setBorderBottom(BorderStyle.THIN);
            generalStyle.setBorderLeft(BorderStyle.THIN);
            generalStyle.setBorderRight(BorderStyle.THIN);

            // 회계 서식
            CellStyle accountingStyle = wb.createCellStyle();
            accountingStyle.setDataFormat((short) 44);
            String csvValue = "";
            boolean isSomethingWrongFile = false;

            for (int i = 2; i < rows.size(); i++) {
                String[] row = rows.get(i);
                Row excelRow = sheet.getRow(startRow);
                if (excelRow == null) {
                    excelRow = sheet.createRow(startRow);
                }

                for (int j = 0; j < 9; j++) {
                    Cell cell = excelRow.createCell(startCol + j);

                    //row의 크기가 가변적이여서 0으로 처리
                    if (row.length < 9 && j >= row.length) {
                        if(!isSomethingWrongFile){
                            linkedHashSet.add(csvFile.getName());
                        }
                        csvValue = "0";
                    } else {
                        csvValue = row[j];
                    }
                    String val = csvValue.trim();

                    try {
                        cell.setCellValue(Double.parseDouble(val.replace(",", "")));
                        cell.setCellStyle(generalStyle);
                    } catch (NumberFormatException e) {
                        cell.setCellValue(val);  // 원본 텍스트 그대로
                        cell.setCellStyle(generalStyle);
                    }
                }
                startRow++;
            }
        }
    }

    public static void writePowerlinkSheet(Sheet sheet, File csvFile, Workbook wb) throws IOException, CsvException {
        System.out.println("파워링크 작업중 csv file name = " + csvFile);
        String encoding = detectEncoding(csvFile);
        try (CSVReader reader = new CSVReader(new InputStreamReader(new FileInputStream(csvFile), Charset.forName(encoding)))) {
            List<String[]> rows = reader.readAll();
            int startRow = 28;  // Excel 29행
            int startCol = 1;   // Excel B열

            DataFormat format = wb.createDataFormat();

            // 일반 서식
            CellStyle generalStyle = wb.createCellStyle();
            generalStyle.setBorderTop(BorderStyle.THIN);
            generalStyle.setBorderBottom(BorderStyle.THIN);
            generalStyle.setBorderLeft(BorderStyle.THIN);
            generalStyle.setBorderRight(BorderStyle.THIN);

            // 회계 서식
            CellStyle accountingStyle = wb.createCellStyle();
            accountingStyle.setDataFormat((short) 44);
            String csvValue = "";
            boolean isSomethingWrongFile = false;

            for (int i = 2; i < rows.size(); i++) {  // 6행부터 시작
                String[] row = rows.get(i);
                Row excelRow = sheet.getRow(startRow);
                if (excelRow == null) {
                    excelRow = sheet.createRow(startRow);
                }

                for (int j = 3; j <= 13; j++) {
                    Cell cell = excelRow.createCell(startCol + (j - 3));

                    //row의 크기가 가변적이여서 0으로 처리
                    if (row.length < 13 && j >= row.length) {

                        if(!isSomethingWrongFile){
                            linkedHashSet.add(csvFile.getName());
                        }
                        csvValue = "0";
                    } else {
                        csvValue = row[j];
                    }
                    String val = csvValue.trim();

                    try {
                        cell.setCellValue(Double.parseDouble(val.replace(",", "")));
                        cell.setCellStyle(generalStyle);
                    } catch (NumberFormatException e) {
                        cell.setCellValue(val);  // 원본 텍스트 그대로
                        cell.setCellStyle(generalStyle);
                    }
                }
                startRow++;
            }
        }
    }

    public static void writeShoppingSheet(Sheet sheet, File csvFile, Workbook wb) throws IOException, CsvException {
        System.out.println("쇼핑시트 작업중 csv file name = " + csvFile);
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
            String csvValue = "";
            boolean isSomethingWrongFile = false;
            CellStyle generalStyle = wb.createCellStyle();
            generalStyle.setBorderTop(BorderStyle.THIN);
            generalStyle.setBorderBottom(BorderStyle.THIN);
            generalStyle.setBorderLeft(BorderStyle.THIN);
            generalStyle.setBorderRight(BorderStyle.THIN);

            for (int i = 2; i < rows.size(); i++) {
                String[] row = rows.get(i);
                if (!"쇼핑검색".equals(row[0])) continue;

                Row excelRow = sheet.getRow(startRow);
                if (excelRow == null) excelRow = sheet.createRow(startRow);

                for (int j = 1; j <= 11; j++) {
                    Cell cell = excelRow.createCell(startCol + (j - 1));

                    //row의 크기가 가변적이여서 0으로 처리
                    if (row.length < 11 && j >= row.length) {

                        if(!isSomethingWrongFile){
                            linkedHashSet.add(csvFile.getName());
                        }
                        csvValue = "0";
                    } else {
                        csvValue = row[j];
                    }
                    String val = csvValue.replace(",", "").trim();

                    try {
                        cell.setCellValue(Double.parseDouble(val.replace(",", "")));
                        cell.setCellStyle(generalStyle);
                    } catch (NumberFormatException e) {
                        cell.setCellValue(val);  // 원본 텍스트 그대로
                        cell.setCellStyle(generalStyle);
                    }

                }
                startRow++;
            }
        }
    }

    public static void writePlaceSheet(Sheet sheet, File csvFile, Workbook wb) throws IOException, CsvException {
        System.out.println("플레이스 작업중 csv file name = " + csvFile);
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
            String csvValue = "";
            boolean isSomethingWrongFile = false;

            DataFormat format = wb.createDataFormat();

            // 일반 서식
            CellStyle generalStyle = wb.createCellStyle();
            generalStyle.setBorderTop(BorderStyle.THIN);
            generalStyle.setBorderBottom(BorderStyle.THIN);
            generalStyle.setBorderLeft(BorderStyle.THIN);
            generalStyle.setBorderRight(BorderStyle.THIN);

            // 회계 서식
            CellStyle accountingStyle = wb.createCellStyle();
            accountingStyle.setDataFormat((short) 44);

            for (int i = 2; i < rows.size(); i++) {
                String[] row = rows.get(i);

                String campaign = row[0].replaceAll("\"", "").trim();
                if (!"플레이스".equals(campaign)) continue;

                Row excelRow = sheet.getRow(startRow);
                if (excelRow == null) excelRow = sheet.createRow(startRow);

                for (int j = 0; j <= 9; j++) {
                    Cell cell = excelRow.createCell(startCol + j);

                    //row의 크기가 가변적이여서 0으로 처리
                    if (row.length < 9 && j >= row.length) {
                        if(!isSomethingWrongFile){
                            linkedHashSet.add(csvFile.getName());
                        }
                        csvValue = "0";
                    } else {
                        csvValue = row[j];
                    }
                    String val = csvValue.replace(",", "").trim();


                    try {
                        double num = Double.parseDouble(val);
                        cell.setCellValue(num);
                        cell.setCellStyle(generalStyle);

                    } catch (NumberFormatException e) {
                        cell.setCellValue(val);
                        cell.setCellStyle(generalStyle);
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

}
