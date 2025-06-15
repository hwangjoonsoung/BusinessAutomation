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
import java.util.*;

public class CsvZipToExcelBatch2 {

    public static void main(String[] args) throws Exception {
        File zipFile = new File("src/main/java/me/hwangjoonsoung/automation/inputCSVZip/place.zip");
        File unzipDir = new File("build/unzipped_place");
        File templateFile = new File("src/main/java/me/hwangjoonsoung/automation/basedExcelFile/12월 키워드보고서_place.xlsx");
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
            File outputFile = new File(outputDir, "12월_키워드보고서_" + id + ".xlsx");

            if (daily.exists()) {
                processOneSet(daily, templatePath, outputFile);
            } else {
                System.out.println("❌ 일별 파일 누락: " + id);
            }
        }
    }

    public static void processOneSet(File dailyCsv, String templatePath, File outputFile) throws Exception {
        FileInputStream fis = new FileInputStream(templatePath);
        Workbook workbook = new XSSFWorkbook(fis);

        Sheet dailySheet = workbook.getSheet("일자별");
        writeDailySheet(dailySheet, dailyCsv, workbook);

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
            int startRow = 28;  // Excel 기준 29행 (AO)
            int startCol = 40;  // Excel 기준 41열 (AO)

            CellStyle numberStyle = wb.createCellStyle();
            DataFormat format = wb.createDataFormat();
            numberStyle.setDataFormat(format.getFormat("#,##0"));

            Font greenFont = wb.createFont();
            greenFont.setColor(IndexedColors.GREEN.getIndex());
            // 🔁 글자 색상을 흰색으로 바꾸고 싶다면 아래처럼 변경하세요:
            // greenFont.setColor(IndexedColors.WHITE.getIndex());

            numberStyle.setFont(greenFont);

            for (int i = 2; i < rows.size(); i++) { // 3행부터 데이터 시작
                String[] row = rows.get(i);
                Row excelRow = sheet.getRow(startRow);
                if (excelRow == null) {
                    excelRow = sheet.createRow(startRow);
                }
                startRow++;

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
