package me.hwangjoonsoung.automation;

import com.opencsv.CSVReader;
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
public class CsvZipToExcelBatch4 {

    public static void main(String[] args) throws Exception {
        File zipFile = new File("src/main/java/me/hwangjoonsoung/automation/inputCSVZip/place.zip");
        File unzipDir = new File("build/unzipped_place");
        File templateFile = new File("src/main/java/me/hwangjoonsoung/automation/basedExcelFile/12월 키워드보고서.xlsx");
        File outputDir = new File("build/output_place");

        if (!outputDir.exists()) outputDir.mkdirs();

        extractZip(zipFile, unzipDir);
        processAllCsvSet(unzipDir, templateFile.getAbsolutePath(), outputDir);
    }

    public static void extractZip(File zipFile, File destDir) throws IOException {
        try (ZipArchiveInputStream zis = new ZipArchiveInputStream(
                new FileInputStream(zipFile), "EUC-KR", false, true)) {
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
        FileInputStream fis = new FileInputStream(templatePath);
        Workbook workbook = new XSSFWorkbook(fis);

        Sheet dailySheet = workbook.getSheet("일자별");
        writeDailySheet(dailySheet, dailyCsv, workbook);

        if (timeCsv.exists()) {
            Sheet timeSheet = workbook.getSheet("시간별");
            writeTimeSheet(timeSheet, timeCsv, workbook);
        }

        if (powerlinkCsv.exists()) {
            Sheet powerlinkSheet = workbook.getSheet("파워링크");
            writePowerlinkSheet(powerlinkSheet, powerlinkCsv, workbook);
        }

        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
            workbook.write(fos);
        }

        if (powerlinkCsv.exists()) {
            Sheet powerlinkSheet = workbook.getSheet("파워링크");
            writePowerlinkSheet(powerlinkSheet, powerlinkCsv, workbook);
        }

        workbook.close();
        System.out.println("✅ 저장 완료: " + outputFile.getAbsolutePath());
    }

    public static void writeTimeSheet(Sheet sheet, File csvFile, Workbook wb) throws IOException, CsvException {
        String encoding = detectEncoding(csvFile);
        try (CSVReader reader = new CSVReader(
                new InputStreamReader(new FileInputStream(csvFile), Charset.forName(encoding)))) {

            List<String[]> rows = reader.readAll();
            int startRow = 60;  // Excel 기준 61행 (B열부터 시작)
            int startCol = 1;   // Excel B열

            DataFormat format = wb.createDataFormat();

            CellStyle defaultStyle = wb.createCellStyle();
            defaultStyle.setDataFormat(format.getFormat("#,##0"));
            Font greenFont = wb.createFont();
            greenFont.setColor(IndexedColors.GREEN.getIndex());
            defaultStyle.setFont(greenFont);

            CellStyle percentStyle = wb.createCellStyle();
            percentStyle.setDataFormat(format.getFormat("0.00%"));
            percentStyle.setFont(greenFont);

            CellStyle floatStyle1 = wb.createCellStyle();
            floatStyle1.setDataFormat(format.getFormat("0.0"));
            floatStyle1.setFont(greenFont);

            CellStyle floatStyle2 = wb.createCellStyle();
            floatStyle2.setDataFormat(format.getFormat("#,##0.##"));
            floatStyle2.setFont(greenFont);

            for (int i = 2; i < rows.size(); i++) {
                String[] row = rows.get(i);
                Row excelRow = sheet.getRow(startRow);
                if (excelRow == null) {
                    excelRow = sheet.createRow(startRow);
                }

                for (int j = 0; j < 9; j++) {
                    Cell cell = excelRow.createCell(startCol + j);
                    String val = row[j].replace(",", "").trim();
                    try {
                        double num = Double.parseDouble(val);
                        cell.setCellValue(num);
                        // 열 인덱스에 따라 다른 스타일 적용
                        if (j == 4) {
                            cell.setCellStyle(floatStyle1); // 평균노출순위
                        } else if (j == 5 || j == 6) {
                            cell.setCellStyle(floatStyle2); // 평균클릭비용, 총비용
                        } else {
                            cell.setCellStyle(defaultStyle);
                        }
                    } catch (NumberFormatException e) {
                        cell.setCellValue(val);
                        cell.setCellStyle(defaultStyle);
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
            CellStyle defaultStyle = wb.createCellStyle();
            defaultStyle.setDataFormat(format.getFormat("#,##0"));
            Font greenFont = wb.createFont();
            greenFont.setColor(IndexedColors.GREEN.getIndex());
            defaultStyle.setFont(greenFont);

            CellStyle floatStyle1 = wb.createCellStyle(); // 평균노출순위
            floatStyle1.setDataFormat(format.getFormat("0.0"));
            floatStyle1.setFont(greenFont);

            CellStyle floatStyle2 = wb.createCellStyle(); // 클릭률/클릭비용 등
            floatStyle2.setDataFormat(format.getFormat("#,##0.00"));
            floatStyle2.setFont(greenFont);

            for (int i = 2; i < rows.size(); i++) {  // 6행부터 시작
                String[] row = rows.get(i);
                Row excelRow = sheet.getRow(startRow);
                if (excelRow == null) excelRow = sheet.createRow(startRow);

                for (int j = 3; j <= 12; j++) {
                    Cell cell = excelRow.createCell(startCol + (j - 3));
                    String val = row[j].replace(",", "").trim();
                    try {
                        double num = Double.parseDouble(val);
                        cell.setCellValue(num);
                        if (j == 6 || j == 7) {
                            cell.setCellStyle(floatStyle2);  // 클릭률, 클릭비용
                        } else if (j == 8) {
                            cell.setCellStyle(floatStyle2);  // 총비용
                        } else if (j == 9) {
                            cell.setCellStyle(floatStyle1);  // 평균노출순위
                        } else {
                            cell.setCellStyle(defaultStyle);
                        }
                    } catch (NumberFormatException e) {
                        cell.setCellValue(val);
                        cell.setCellStyle(defaultStyle);
                    }
                }
                startRow++;
            }
        }
    }

    public static void writeDailySheet(Sheet sheet, File csvFile, Workbook wb) throws IOException, CsvException {
        String encoding = detectEncoding(csvFile);
        try (CSVReader reader = new CSVReader(
                new InputStreamReader(new FileInputStream(csvFile), Charset.forName(encoding)))) {

            List<String[]> rows = reader.readAll();
            int startRow = 28;  // Excel 기준 29행 (AO)
            int startCol = 40;  // Excel 기준 41열 (AO)

            CellStyle numberStyle = wb.createCellStyle();
            DataFormat format = wb.createDataFormat();
            numberStyle.setDataFormat(format.getFormat("#,##0"));

            Font greenFont = wb.createFont();
            greenFont.setColor(IndexedColors.GREEN.getIndex());
            numberStyle.setFont(greenFont);

            for (int i = 2; i < rows.size(); i++) { // 3행부터 데이터 시작
                String[] row = rows.get(i);
                Row excelRow = sheet.getRow(startRow);
                if (excelRow == null) {
                    excelRow = sheet.createRow(startRow);
                } else {
                    for (int c = startCol; c < startCol + 9; c++) {
                        Cell cell = excelRow.getCell(c);
                        if (cell != null) cell.setBlank();
                    }
                }

                for (int j = 0; j < 9; j++) {
                    Cell cell = excelRow.createCell(startCol + j);
                    String val = row[j].replace(",", "").trim();

                    try {
                        double num = Double.parseDouble(val);
                        cell.setCellValue(num);
                        cell.setCellStyle(numberStyle);
                    } catch (NumberFormatException e) {
                        cell.setCellValue(val);
                        cell.setCellStyle(numberStyle);
                    }
                }
                startRow++;
            }

            int cleanupStartRow = startRow;
            int maxRows = sheet.getLastRowNum();
            for (int i = cleanupStartRow; i <= maxRows; i++) {
                Row row = sheet.getRow(i);
                if (row != null) sheet.removeRow(row);
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
