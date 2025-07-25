import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
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
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

public class CsvZipToExcelBatch2 {

    static LinkedHashSet linkedHashSet = new LinkedHashSet();
    public static void main(String[] args) throws Exception {
        File zipFile = new File("src/main/java/inputCSVZip/csv.zip");
        File unzipDir = new File("src/main/java/unzipped");
        File templateFile = new File("src/main/java/basedExcelFile/06월 키워드보고서.xlsx");
        File outputDir = new File("src/main/java/output");

        if (!outputDir.exists()) outputDir.mkdirs();

        extractZip(zipFile, unzipDir);
        processAllCsvSet(unzipDir, templateFile.getAbsolutePath(), outputDir);

        System.out.println("===========================================\n뭔가 이상한 파일들 : " + linkedHashSet);

        // 파일 자동 지워짐
        File[] files = unzipDir.listFiles();
        for (File file : files) {
            file.delete();
        }
        boolean delete = unzipDir.delete();
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
            File outputFile = new File(outputDir, "06월_키워드보고서_" + id + ".xlsx");

//            if(!id.equals("kjss2106_naver")){
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
        String baseName = outputFile.getName().replace("06월_키워드보고서_", "").replace(".xlsx", "");
        File powerlinkCsv = new File("src/main/java/unzipped/파워링크보고서," + baseName + ".csv");
        File shoppingCsv = new File("src/main/java/unzipped/쇼핑검색보고서," + baseName + ".csv");
        File placeCsv = new File("src/main/java/unzipped/플레이스보고서," + baseName + ".csv");
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


        // ✅ 추가된 핵심 한 줄: 수식 강제 재계산 설정
        workbook.setForceFormulaRecalculation(true);

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

                    try {
                        if (row.length < 9 && j >= row.length) {
                            if (!isSomethingWrongFile) {
                                linkedHashSet.add(csvFile.getName());
                            }
                            csvValue = "0";
                        } else {
                            csvValue = row[j];
                        }
                    } catch (ArrayIndexOutOfBoundsException e) {
                        linkedHashSet.add(csvFile.getName());
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
            generalStyle.setDataFormat(format.getFormat("#,##0"));
            generalStyle.setBorderTop(BorderStyle.THIN);
            generalStyle.setBorderBottom(BorderStyle.THIN);
            generalStyle.setBorderLeft(BorderStyle.THIN);
            generalStyle.setBorderRight(BorderStyle.THIN);

            CellStyle generalStyle2 = wb.createCellStyle();
            generalStyle2.setBorderTop(BorderStyle.THIN);
            generalStyle2.setBorderBottom(BorderStyle.THIN);
            generalStyle2.setBorderLeft(BorderStyle.THIN);
            generalStyle2.setBorderRight(BorderStyle.THIN);

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
                    try {
                        //row의 크기가 가변적이여서 0으로 처리
                        if (row.length < 9 && j >= row.length) {
                            if (!isSomethingWrongFile) {
                                linkedHashSet.add(csvFile.getName());
                            }
                            csvValue = "0";
                        } else {
                            csvValue = row[j];
                        }
                    } catch (ArrayIndexOutOfBoundsException e) {
                        linkedHashSet.add(csvFile.getName());
                    }
                    String val = csvValue.trim();

                    try {
                        cell.setCellValue(Double.parseDouble(val.replace(",", "")));
                        // todo: 일반으로 하는건 그냥 두고 자리수 표시 해야하는걸 else로 적용
                        /*
                        * D : 2
                        * E : 3
                        * F : 4
                        * G : 5
                        * H : 6
                        * I : 7
                        * J : 8
                        * */
                        if(j == 3 || j == 4){
                            cell.setCellStyle(generalStyle2);
                        }else{
                            cell.setCellStyle(generalStyle);
                        }
                    } catch (NumberFormatException e) {
                        cell.setCellValue(val);  // 원본 텍스트 그대로
                        if(j == 3 || j == 4){
                            cell.setCellStyle(generalStyle2);
                        }else{
                            cell.setCellStyle(generalStyle);
                        }
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
            generalStyle.setDataFormat(format.getFormat("#,##0"));
            generalStyle.setBorderTop(BorderStyle.THIN);
            generalStyle.setBorderBottom(BorderStyle.THIN);
            generalStyle.setBorderLeft(BorderStyle.THIN);
            generalStyle.setBorderRight(BorderStyle.THIN);

            CellStyle generalStyle2 = wb.createCellStyle();
            generalStyle2.setBorderTop(BorderStyle.THIN);
            generalStyle2.setBorderBottom(BorderStyle.THIN);
            generalStyle2.setBorderLeft(BorderStyle.THIN);
            generalStyle2.setBorderRight(BorderStyle.THIN);

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

                    try {
                        //row의 크기가 가변적이여서 0으로 처리
                        if (row.length < 13 && j >= row.length) {
                            if (!isSomethingWrongFile) {
                                linkedHashSet.add(csvFile.getName());
                            }
                            csvValue = "0";
                        } else {
                            csvValue = row[j];
                        }
                    } catch (ArrayIndexOutOfBoundsException e) {
                        linkedHashSet.add(csvFile.getName());
                    }
                    String val = csvValue.trim();

                    try {
                        cell.setCellValue(Double.parseDouble(val.replace(",", "")));
                        if(j == 9 || j == 10 || j ==11 ){
                            cell.setCellStyle(generalStyle2);
                        }else{
                            cell.setCellStyle(generalStyle);
                        }
                    } catch (NumberFormatException e) {
                        cell.setCellValue(val);  // 원본 텍스트 그대로
                        if(j == 9 || j == 10 || j ==11 ){
                            cell.setCellStyle(generalStyle2);
                        }else{
                            cell.setCellStyle(generalStyle);
                        }
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
            DataFormat format = wb.createDataFormat();
            boolean isSomethingWrongFile = false;

            CellStyle generalStyle = wb.createCellStyle();
            generalStyle.setDataFormat(format.getFormat("#,##0"));
            generalStyle.setBorderTop(BorderStyle.THIN);
            generalStyle.setBorderBottom(BorderStyle.THIN);
            generalStyle.setBorderLeft(BorderStyle.THIN);
            generalStyle.setBorderRight(BorderStyle.THIN);

            CellStyle generalStyle2 = wb.createCellStyle();
            generalStyle2.setBorderTop(BorderStyle.THIN);
            generalStyle2.setBorderBottom(BorderStyle.THIN);
            generalStyle2.setBorderLeft(BorderStyle.THIN);
            generalStyle2.setBorderRight(BorderStyle.THIN);

            for (int i = 2; i < rows.size(); i++) {
                String[] row = rows.get(i);
                if (!"쇼핑검색".equals(row[0])) continue;

                Row excelRow = sheet.getRow(startRow);
                if (excelRow == null) excelRow = sheet.createRow(startRow);

                for (int j = 1; j <= 11; j++) {
                    Cell cell = excelRow.createCell(startCol + (j - 1));

                    try {
                        //row의 크기가 가변적이여서 0으로 처리
                        if (row.length < 11 && j >= row.length) {
                            if (!isSomethingWrongFile) {
                                linkedHashSet.add(csvFile.getName());
                            }
                            csvValue = "0";
                        } else {
                            csvValue = row[j];
                        }
                    } catch (ArrayIndexOutOfBoundsException e) {
                        linkedHashSet.add(csvFile.getName());
                    }
                    String val = csvValue.replace(",", "").trim();

                    try {
                        cell.setCellValue(Double.parseDouble(val.replace(",", "")));
                        if(j == 7 || j == 8 || j ==9 ){
                            cell.setCellStyle(generalStyle2);
                        }else{
                            cell.setCellStyle(generalStyle);
                        }
                    } catch (NumberFormatException e) {
                        cell.setCellValue(val);  // 원본 텍스트 그대로
                        if(j == 7 || j == 8 || j ==9 ){
                            cell.setCellStyle(generalStyle2);
                        }else{
                            cell.setCellStyle(generalStyle);
                        }
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
            generalStyle.setDataFormat(format.getFormat("#,##0"));
            generalStyle.setBorderTop(BorderStyle.THIN);
            generalStyle.setBorderBottom(BorderStyle.THIN);
            generalStyle.setBorderLeft(BorderStyle.THIN);
            generalStyle.setBorderRight(BorderStyle.THIN);

            // 일반 서식
            CellStyle generalStyle2 = wb.createCellStyle();
            generalStyle2.setBorderTop(BorderStyle.THIN);
            generalStyle2.setBorderBottom(BorderStyle.THIN);
            generalStyle2.setBorderLeft(BorderStyle.THIN);
            generalStyle2.setBorderRight(BorderStyle.THIN);

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

                    try {
                        //row의 크기가 가변적이여서 0으로 처리
                        if (row.length < 9 && j >= row.length) {
                            if (!isSomethingWrongFile) {
                                linkedHashSet.add(csvFile.getName());
                            }
                            csvValue = "0";
                        } else {
                            csvValue = row[j];
                        }
                    } catch (ArrayIndexOutOfBoundsException e) {
                        linkedHashSet.add(csvFile.getName());
                    }
                    String val = csvValue.replace(",", "").trim();


                    try {
                        double num = Double.parseDouble(val);
                        cell.setCellValue(num);
                        /*
                         * F : 4
                         * G : 5
                         * H : 6
                         * I : 7
                         * J : 8
                         * K : 9
                         * */
                        if(j == 9){
                            cell.setCellStyle(generalStyle2);
                        }else{
                            cell.setCellStyle(generalStyle);
                        }
                    } catch (NumberFormatException e) {
                        cell.setCellValue(val);
                        if(j == 9){
                            cell.setCellStyle(generalStyle2);
                        }else{
                            cell.setCellStyle(generalStyle);
                        }
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

    public static void runVbaAndSaveAsXlsxBatch(File inputDir, File outputDir) {
        if (!inputDir.exists() || !inputDir.isDirectory()) {
            System.err.println("❌ 입력 폴더가 존재하지 않습니다: " + inputDir.getAbsolutePath());
            return;
        }

        if (!outputDir.exists()) {
            outputDir.mkdirs();
        }

        ActiveXComponent excel = new ActiveXComponent("Excel.Application");

        try {
            excel.setProperty("Visible", false);
            Dispatch workbooks = excel.getProperty("Workbooks").toDispatch();

            File[] xlsxFiles = inputDir.listFiles((dir, name) -> name.endsWith(".xlsx"));
            if (xlsxFiles == null || xlsxFiles.length == 0) {
                System.out.println("⚠️ .xlsx 파일 없음: " + inputDir.getAbsolutePath());
                return;
            }

            for (File xlsxFile : xlsxFiles) {
                String fileName = xlsxFile.getName();
                System.out.println("▶ 처리 중: " + fileName);

                try {
                    // 1. 열기
                    Dispatch workbook = Dispatch.call(workbooks, "Open", xlsxFile.getAbsolutePath()).toDispatch();

                    // 2. VBA 매크로 실행
                    Dispatch.call(excel, "Run", "RunMacroManually");  // ❗ 여기에 Sub 이름

                    // 3. .xlsx로 저장 (형식 코드 51 = xlOpenXMLWorkbook)
                    String outputName = fileName.replace(".xlsx", ".xlsx");
                    File outputFile = new File(outputDir, outputName);
                    Dispatch.call(workbook, "SaveAs", outputFile.getAbsolutePath(), 51);

                    // 4. 닫기
                    Dispatch.call(workbook, "Close", false);

                    System.out.println("✅ 저장 완료: " + outputFile.getAbsolutePath());

                } catch (Exception e) {
                    System.err.println("❌ 처리 실패: " + fileName);
                    e.printStackTrace();
                }
            }
        } finally {
            excel.invoke("Quit");
        }
    }
}
