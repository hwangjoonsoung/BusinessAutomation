package me.hwangjoonsoung.automation;

import com.opencsv.CSVReader;
import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipArchiveInputStream;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

public class CsvZipToExcelBatch {

    public static void main(String[] args) throws Exception {
        // 상대 경로 기준 설정 (루트: 프로젝트 디렉토리)
        File zipFile = new File("src/main/java/me/hwangjoonsoung/automation/inputCSVZip/place.zip");
        File unzipDir = new File("build/unzipped_place");  // 출력 디렉토리는 build 아래로
        File templateFile = new File("src/main/java/me/hwangjoonsoung/automation/basedExcelFile/12월 키워드보고서_place.xlsx");
        File outputDir = new File("build/output_place");

        if (!outputDir.exists()) outputDir.mkdirs();

        CsvZipToExcelBatch.extractZip(zipFile, unzipDir);
        CsvZipToExcelBatch.processAllCsvSet(unzipDir, templateFile.getAbsolutePath(), outputDir);
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

        // 파일 이름 기반으로 ID 추출 및 처리
        Set<String> idSet = new HashSet<>();
        for (File f : files) {
            String name = f.getName();
            if (name.startsWith("파워링크보고서,")) {
                String id = name.replace("파워링크보고서,", "").replace(".csv", "");
                idSet.add(id);
            }
        }

        for (String id : idSet) {
            File main = new File(folder, "파워링크보고서," + id + ".csv");
            File daily = new File(folder, "일별보고서," + id + ".csv");
            File time = new File(folder, "요일별보고서," + id + ".csv");

            if (main.exists() && daily.exists() && time.exists()) {
                File outputFile = new File(outputDir, "12월_키워드보고서_" + id + ".xlsx");
                processOneSet(main, daily, time, templatePath, outputFile);
            } else {
                System.out.println("❌ 일부 파일 누락: " + id);
            }
        }
    }

    public static void processOneSet(File mainCsv, File dailyCsv, File timeCsv, String templatePath, File outputFile) throws Exception {
        FileInputStream fis = new FileInputStream(templatePath);
        Workbook workbook = new XSSFWorkbook(fis);

        Sheet sheet = workbook.getSheet("파워링크");
        try (CSVReader reader = new CSVReader(new InputStreamReader(new FileInputStream(mainCsv), StandardCharsets.UTF_8))) {
            List<String[]> rows = reader.readAll();
            int excelStartRow = 28; // 0-based → 실제 Excel은 29행

            for (int i = 5; i < rows.size(); i++) { // 헤더 스킵
                String[] row = rows.get(i);
                Row excelRow = sheet.createRow(excelStartRow++);
                for (int j = 3; j < row.length; j++) { // 3열 이후부터 복사
                    Cell cell = excelRow.createCell(j - 3);
                    try {
                        cell.setCellValue(Double.parseDouble(row[j].replace(",", "")));
                    } catch (Exception e) {
                        cell.setCellValue(row[j]);
                    }
                }
            }
        }

        // TODO: 일자별(dailyCsv), 요일별(timeCsv) 도 동일하게 추가 가능

        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
            workbook.write(fos);
        }

        workbook.close();
        System.out.println("✅ 저장 완료: " + outputFile.getAbsolutePath());
    }
}
