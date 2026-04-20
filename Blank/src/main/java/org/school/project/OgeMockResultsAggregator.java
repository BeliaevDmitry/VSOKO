package org.school.project;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.*;
import java.sql.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.stream.Collectors;

/**
 * Агрегатор результатов пробников ОГЭ.
 *
 * <p>Что делает программа:</p>
 * <ol>
 *     <li>Читает выбор экзаменов из файла админа (лист "Выбор экзамена").</li>
 *     <li>Читает шкалы перевода баллов в оценку с листа "Баллы".</li>
 *     <li>Сканирует папку с отчётами пробников (xlsx, листы "Информация" и "Сбор информации").</li>
 *     <li>Сохраняет/обновляет данные во временной БД H2 (M2).</li>
 *     <li>Формирует сводный excel-файл с 3 листами.</li>
 * </ol>
 */
public class OgeMockResultsAggregator {
    private static final Logger LOG = Logger.getLogger(OgeMockResultsAggregator.class.getName());

    private static final String REPORTS_FOLDER =
            "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ №7\\ВСОКО\\Работы 2025 2026\\ОГЭ\\апрель 2025";
    private static final String ADMIN_FILE =
            "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ №7\\ОГЭ 2026\\ОГЭ 2026 админ.xlsx";

    private static final String DB_FILE =
            "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ №7\\ВСОКО\\Работы 2025 2026\\ОГЭ\\апрель 2025\\m2_oge_results";

    private static final List<String> SUBJECT_ORDER = List.of(
            "Русский язык", "Математика", "Физика", "Химия", "Информатика и ИКТ", "Биология",
            "История", "География", "Английский язык", "Немецкий язык", "Французский язык",
            "Обществознание", "Испанский язык", "Литература"
    );

    private static final Map<String, String> SUBJECT_ALIASES = new LinkedHashMap<>();

    static {
        SUBJECT_ALIASES.put("Информатика", "Информатика и ИКТ");
        SUBJECT_ALIASES.put("Информатика и ИКТ", "Информатика и ИКТ");
        SUBJECT_ALIASES.put("Русский", "Русский язык");
        SUBJECT_ALIASES.put("Матем", "Математика");
        SUBJECT_ALIASES.put("Английский", "Английский язык");
        SUBJECT_ALIASES.put("Немецкий", "Немецкий язык");
        SUBJECT_ALIASES.put("Французский", "Французский язык");
        SUBJECT_ALIASES.put("Испанский", "Испанский язык");
        SUBJECT_ALIASES.put("Общество", "Обществознание");
    }

    public static void main(String[] args) {
        try {
            LOG.info("Старт OGE Mock Aggregator");
            LOG.info("Папка отчетов: " + REPORTS_FOLDER);
            Path outputPath = Paths.get(REPORTS_FOLDER,
                    "Свод_пробники_ОГЭ_" + LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmm")) + ".xlsx");

            try (Connection connection = DriverManager.getConnection("jdbc:h2:file:" + DB_FILE + ";AUTO_SERVER=TRUE");
                 Workbook adminBook = new XSSFWorkbook(new FileInputStream(ADMIN_FILE))) {

                initDb(connection);
                loadExamChoices(connection, adminBook.getSheet("Выбор экзамена"));
                Map<String, Map<Integer, Integer>> scoreToGrade = loadScoreToGrade(adminBook.getSheet("Баллы"));
                loadMockReports(connection, scoreToGrade);
                generateOutputWorkbook(connection, outputPath);

                LOG.info("Готово. Файл отчета: " + outputPath);
            }
        } catch (Exception e) {
            LOG.log(Level.SEVERE, "Критическая ошибка выполнения", e);
            throw new RuntimeException(e);
        }
    }

    private static void initDb(Connection connection) throws SQLException {
        try (Statement st = connection.createStatement()) {
            st.execute("""
                    CREATE TABLE IF NOT EXISTS exam_choices (
                        class_name VARCHAR(100),
                        fio VARCHAR(500) NOT NULL,
                        subject VARCHAR(200) NOT NULL,
                        PRIMARY KEY (fio, subject)
                    )
                    """);
            try {
                st.execute("ALTER TABLE exam_choices ADD COLUMN IF NOT EXISTS class_name VARCHAR(100)");
            } catch (SQLException ignored) {
                // На старых версиях H2 IF NOT EXISTS может не поддерживаться в ALTER, оставляем совместимость.
            }
            st.execute("""
                    CREATE TABLE IF NOT EXISTS mock_results (
                        fio VARCHAR(500) NOT NULL,
                        class_name VARCHAR(100),
                        subject VARCHAR(200) NOT NULL,
                        test_score INT,
                        grade INT,
                        source_file VARCHAR(1000),
                        loaded_at TIMESTAMP,
                        PRIMARY KEY (fio, subject)
                    )
                    """);
        }
    }

    private static void loadExamChoices(Connection connection, Sheet examSheet) throws SQLException {
        if (examSheet == null) {
            throw new IllegalArgumentException("Лист 'Выбор экзамена' не найден в файле админа.");
        }
        LOG.info("Загрузка выборов экзаменов из листа 'Выбор экзамена'");

        try (PreparedStatement upsert = connection.prepareStatement("""
                MERGE INTO exam_choices (class_name, fio, subject) KEY (fio, subject) VALUES (?, ?, ?)
                """)) {

            for (int rowIdx = 1; rowIdx <= examSheet.getLastRowNum(); rowIdx++) {
                Row row = examSheet.getRow(rowIdx);
                if (row == null) continue;

                String className = normalizeClassName(getCellText(row.getCell(0)).trim());
                String fio = getCellText(row.getCell(1)).trim();
                String subjectsRaw = getCellText(row.getCell(3)).trim();
                if (fio.isEmpty() || subjectsRaw.isEmpty()) continue;

                Set<String> subjects = extractSubjects(subjectsRaw);
                for (String subject : subjects) {
                    upsert.setString(1, className);
                    upsert.setString(2, normalizeFio(fio));
                    upsert.setString(3, subject);
                    upsert.addBatch();
                }
            }
            upsert.executeBatch();
        }
        LOG.info("Загрузка выборов завершена");
    }

    private static Map<String, Map<Integer, Integer>> loadScoreToGrade(Sheet scoreSheet) {
        if (scoreSheet == null) {
            throw new IllegalArgumentException("Лист 'Баллы' не найден в файле админа.");
        }

        Map<String, Integer> subjectCol = new LinkedHashMap<>();
        Row header = scoreSheet.getRow(0);
        if (header == null) {
            throw new IllegalArgumentException("Лист 'Баллы' пустой.");
        }

        for (int c = 1; c <= 40; c++) {
            Cell cell = header.getCell(c);
            String raw = getCellText(cell);
            String normalized = normalizeSubject(raw);
            if (normalized != null) {
                subjectCol.put(normalized, c);
            }
        }

        Map<String, Map<Integer, Integer>> map = new LinkedHashMap<>();
        for (String subject : SUBJECT_ORDER) {
            map.put(subject, new HashMap<>());
        }

        for (int r = 1; r <= scoreSheet.getLastRowNum(); r++) {
            Row row = scoreSheet.getRow(r);
            if (row == null) continue;

            Integer score = parseInt(getCellText(row.getCell(0)));
            if (score == null) continue;

            for (String subject : SUBJECT_ORDER) {
                Integer col = subjectCol.get(subject);
                if (col == null) continue;
                Integer grade = parseInt(getCellText(row.getCell(col)));
                if (grade != null) {
                    map.get(subject).put(score, grade);
                }
            }
        }

        return map;
    }

    private static void loadMockReports(Connection connection,
                                        Map<String, Map<Integer, Integer>> scoreToGrade) throws IOException, SQLException {
        try (Statement st = connection.createStatement()) {
            // Пересобираем результаты из текущего набора файлов, чтобы не тянуть ошибочные
            // значения из прошлых запусков (например, старые "0/2" по пустым строкам).
            st.execute("TRUNCATE TABLE mock_results");
        }

        List<Path> files;
        try (var stream = Files.walk(Paths.get(REPORTS_FOLDER))) {
            files = stream
                    .filter(Files::isRegularFile)
                    .filter(p -> p.toString().toLowerCase(Locale.ROOT).endsWith(".xlsx"))
                    .filter(p -> !p.getFileName().toString().toLowerCase(Locale.ROOT).contains("свод_пробники_огэ"))
                    .filter(p -> !p.getFileName().toString().toLowerCase(Locale.ROOT).contains("админ"))
                    .collect(Collectors.toList());
        }
        LOG.info("Найдено excel-файлов для обработки: " + files.size());

        try (PreparedStatement upsert = connection.prepareStatement("""
                MERGE INTO mock_results (fio, class_name, subject, test_score, grade, source_file, loaded_at)
                KEY (fio, subject)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """)) {

            for (Path file : files) {
                LOG.info("Обработка файла: " + file);
                try (Workbook workbook = new XSSFWorkbook(new FileInputStream(file.toFile()))) {
                    String subject = extractSubject(workbook.getSheet("Информация"), file);
                    if (subject == null || !SUBJECT_ORDER.contains(subject)) {
                        LOG.warning("Не удалось определить предмет, файл пропущен: " + file.getFileName());
                        continue;
                    }

                    Sheet data = workbook.getSheet("Сбор информации");
                    if (data == null) {
                        LOG.warning("Лист 'Сбор информации' отсутствует, файл пропущен: " + file.getFileName());
                        continue;
                    }
                    HeaderPos pos = detectHeader(data);
                    if (pos == null) {
                        LOG.warning("Не найдены нужные заголовки (ФИО/Итог), файл пропущен: " + file.getFileName());
                        continue;
                    }

                    for (int r = pos.dataStartRow; r <= data.getLastRowNum(); r++) {
                        Row row = data.getRow(r);
                        if (row == null) continue;

                        String fio = normalizeFio(getCellText(row.getCell(pos.fioCol)));
                        if (fio.isEmpty()) continue;

                        String classFromRow = pos.classCol >= 0 ? getCellText(row.getCell(pos.classCol)).trim() : "";
                        String classFromFile = extractClassFromFileName(file.getFileName().toString());
                        String className = normalizeClassName(!classFromRow.isBlank() ? classFromRow : classFromFile);
                        Integer score = pos.scoreCol >= 0 ? parseInt(getCellText(row.getCell(pos.scoreCol))) : null;
                        if (!isMeaningfulResultRow(row, pos, score)) {
                            LOG.fine("Пропуск незаполненной строки: " + fio + " (" + file.getFileName() + ")");
                            continue;
                        }
                        Integer grade = score == null ? null : scoreToGrade.getOrDefault(subject, Map.of()).get(score);

                        upsert.setString(1, fio);
                        upsert.setString(2, className);
                        upsert.setString(3, subject);
                        if (score == null) upsert.setNull(4, Types.INTEGER); else upsert.setInt(4, score);
                        if (grade == null) upsert.setNull(5, Types.INTEGER); else upsert.setInt(5, grade);
                        upsert.setString(6, file.toString());
                        upsert.setTimestamp(7, Timestamp.valueOf(LocalDateTime.now()));
                        upsert.addBatch();
                    }
                } catch (Exception ex) {
                    LOG.log(Level.WARNING, "Пропуск файла из-за ошибки: " + file, ex);
                }
            }

            upsert.executeBatch();
        }
        LOG.info("Загрузка результатов из отчетов завершена");
    }

    private static void generateOutputWorkbook(Connection connection, Path outputPath) throws SQLException, IOException {
        Map<String, StudentRow> students = loadStudents(connection);
        List<MissingItem> missingItems = new ArrayList<>();

        try (Workbook workbook = new XSSFWorkbook()) {
            CellStyle headerStyle = createHeaderStyle(workbook);
            CellStyle redStyle = createFillStyle(workbook, IndexedColors.ROSE.getIndex());
            CellStyle yellowStyle = createFillStyle(workbook, IndexedColors.LIGHT_YELLOW.getIndex());

            Sheet results = workbook.createSheet("Результаты пробника");
            createResultsHeader(results, headerStyle);

        List<StudentRow> orderedStudents = students.values().stream()
                .sorted(Comparator
                        .comparing((StudentRow s) -> sortableClass(s.className))
                        .thenComparing(s -> normalizeFio(s.fio), String.CASE_INSENSITIVE_ORDER))
                .toList();

        int rowNum = 2;
        for (StudentRow student : orderedStudents) {
                Row row = results.createRow(rowNum++);
                row.createCell(0).setCellValue(student.className == null ? "" : student.className);
                row.createCell(1).setCellValue(student.fio);

                int col = 2;
                for (String subject : SUBJECT_ORDER) {
                    ResultCell rc = student.results.get(subject);
                    Cell scoreCell = row.createCell(col++);
                    Cell gradeCell = row.createCell(col++);

                    if (rc != null && rc.score != null) scoreCell.setCellValue(rc.score);
                    if (rc != null && rc.grade != null) gradeCell.setCellValue(rc.grade);

                    boolean studentChoosesSubject = student.subjects.contains(subject);
                    boolean hasResult = rc != null && rc.score != null;

                    if (rc != null && rc.grade != null && rc.grade == 2) {
                        scoreCell.setCellStyle(redStyle);
                        gradeCell.setCellStyle(redStyle);
                    }

                    if (studentChoosesSubject && !hasResult) {
                        scoreCell.setCellStyle(yellowStyle);
                        gradeCell.setCellStyle(yellowStyle);
                        missingItems.add(new MissingItem(subject, student.fio));
                    }
                }
            }

            autoSize(results, 2 + SUBJECT_ORDER.size() * 2);
            results.createFreezePane(0, 1);

            Sheet missing = workbook.createSheet("Недостающая информация");
            Row mh = missing.createRow(0);
            mh.createCell(0).setCellValue("Предмет");
            mh.createCell(1).setCellValue("ФИО");
            mh.getCell(0).setCellStyle(headerStyle);
            mh.getCell(1).setCellStyle(headerStyle);

            int mr = 1;
            for (MissingItem item : missingItems) {
                Row row = missing.createRow(mr++);
                row.createCell(0).setCellValue(item.subject);
                row.createCell(1).setCellValue(item.fio);
            }
            autoSize(missing, 2);

            Sheet stats = workbook.createSheet("Статистика");
            generateStatsSheet(stats, students, headerStyle);
            autoSize(stats, 8);

            Files.createDirectories(outputPath.getParent());
            try (FileOutputStream fos = new FileOutputStream(outputPath.toFile())) {
                workbook.write(fos);
            }
        }
    }

    private static Map<String, StudentRow> loadStudents(Connection connection) throws SQLException {
        Map<String, StudentRow> map = new TreeMap<>();

        try (Statement st = connection.createStatement();
             ResultSet rs = st.executeQuery("SELECT class_name, fio, subject FROM exam_choices")) {
            while (rs.next()) {
                String className = normalizeClassName(rs.getString("class_name"));
                String fio = rs.getString("fio");
                String subject = rs.getString("subject");
                StudentRow student = map.computeIfAbsent(fio, k -> new StudentRow(fio));
                if (student.className == null || student.className.isBlank()) student.className = className;
                student.subjects.add(subject);
            }
        }

        try (Statement st = connection.createStatement();
             ResultSet rs = st.executeQuery("SELECT fio, class_name, subject, test_score, grade FROM mock_results")) {
            while (rs.next()) {
                String fio = rs.getString("fio");
                StudentRow row = map.computeIfAbsent(fio, k -> new StudentRow(fio));
                String className = rs.getString("class_name");
                if (row.className == null || row.className.isBlank()) row.className = normalizeClassName(className);

                String subject = rs.getString("subject");
                Integer score = (Integer) rs.getObject("test_score");
                Integer grade = (Integer) rs.getObject("grade");
                row.results.put(subject, new ResultCell(score, grade));
            }
        }

        return map;
    }

    private static void generateStatsSheet(Sheet stats, Map<String, StudentRow> students, CellStyle headerStyle) {
        Row header = stats.createRow(0);
        String[] cols = {"Класс", "Предмет", "Двоек", "Троек", "Четверок", "Пятерок"};
        for (int i = 0; i < cols.length; i++) {
            header.createCell(i).setCellValue(cols[i]);
            header.getCell(i).setCellStyle(headerStyle);
        }

        Map<String, Map<String, int[]>> agg = new TreeMap<>();
        for (StudentRow student : students.values()) {
            String cls = (student.className == null || student.className.isBlank()) ? "Не указан" : student.className;
            for (Map.Entry<String, ResultCell> e : student.results.entrySet()) {
                Integer g = e.getValue().grade;
                if (g == null || g < 2 || g > 5) continue;
                agg.computeIfAbsent(cls, k -> new TreeMap<>())
                        .computeIfAbsent(e.getKey(), k -> new int[6])[g]++;
            }
        }

        int r = 1;
        for (var clsEntry : agg.entrySet()) {
            for (var subEntry : clsEntry.getValue().entrySet()) {
                int[] a = subEntry.getValue();
                Row row = stats.createRow(r++);
                row.createCell(0).setCellValue(clsEntry.getKey());
                row.createCell(1).setCellValue(subEntry.getKey());
                row.createCell(2).setCellValue(a[2]);
                row.createCell(3).setCellValue(a[3]);
                row.createCell(4).setCellValue(a[4]);
                row.createCell(5).setCellValue(a[5]);
            }
        }

        r += 1;
        Row blockTitle = stats.createRow(r++);
        blockTitle.createCell(0).setCellValue("Ученики с 2+ двойками");
        blockTitle.getCell(0).setCellStyle(headerStyle);

        Row h2 = stats.createRow(r++);
        h2.createCell(0).setCellValue("Класс");
        h2.createCell(1).setCellValue("ФИО");
        h2.createCell(2).setCellValue("Количество двоек");
        h2.createCell(3).setCellValue("Предметы (оценка 2)");
        h2.getCell(0).setCellStyle(headerStyle);
        h2.getCell(1).setCellStyle(headerStyle);
        h2.getCell(2).setCellStyle(headerStyle);
        h2.getCell(3).setCellStyle(headerStyle);

        for (StudentRow student : students.values()) {
            List<String> twoSubjects = student.results.entrySet().stream()
                    .filter(e -> e.getValue().grade != null && e.getValue().grade == 2)
                    .map(Map.Entry::getKey)
                    .sorted(String.CASE_INSENSITIVE_ORDER)
                    .toList();
            long twos = twoSubjects.size();
            if (twos >= 2) {
                Row row = stats.createRow(r++);
                row.createCell(0).setCellValue(student.className == null ? "" : student.className);
                row.createCell(1).setCellValue(student.fio);
                row.createCell(2).setCellValue(twos);
                row.createCell(3).setCellValue(String.join(", ", twoSubjects));
            }
        }
    }

    private static void createResultsHeader(Sheet sheet, CellStyle headerStyle) {
        Row row0 = sheet.createRow(0);
        Row row1 = sheet.createRow(1);

        Cell c0 = row0.createCell(0);
        c0.setCellValue("Класс");
        c0.setCellStyle(headerStyle);
        Cell c1 = row0.createCell(1);
        c1.setCellValue("ФИО");
        c1.setCellStyle(headerStyle);

        row1.createCell(0).setCellStyle(headerStyle);
        row1.createCell(1).setCellStyle(headerStyle);

        int col = 2;
        for (String subject : SUBJECT_ORDER) {
            Cell head = row0.createCell(col);
            head.setCellValue(subject);
            head.setCellStyle(headerStyle);

            Cell sub1 = row1.createCell(col);
            sub1.setCellValue("Тестовый балл");
            sub1.setCellStyle(headerStyle);

            Cell sub2 = row1.createCell(col + 1);
            sub2.setCellValue("Оценка");
            sub2.setCellStyle(headerStyle);

            sheet.addMergedRegion(new CellRangeAddress(0, 0, col, col + 1));
            col += 2;
        }

        sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 1, 1));
    }

    private static HeaderPos detectHeader(Sheet sheet) {
        // Бизнес-правило для текущих форматов:
        // 1-я строка содержит заголовки с ФИО/Класс/Итог (или Итого),
        // данные начинаются с 4-й строки (после 3 строк шапки).
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return null;

        int fioCol = -1;
        int classCol = -1;
        int presenceCol = -1;
        int variantCol = -1;
        int scoreCol = -1;

        for (int c = 0; c <= Math.min(70, headerRow.getLastCellNum()); c++) {
            String val = getCellText(headerRow.getCell(c)).toLowerCase(Locale.ROOT);
            if (val.contains("фио")) fioCol = c;
            if (val.contains("класс")) classCol = c;
            if (val.contains("присутств")) presenceCol = c;
            if (val.contains("вариант")) variantCol = c;
            String compact = val.replaceAll("\\s+", "");
            if ("итог".equals(compact) || "итого".equals(compact)) scoreCol = c;
        }

        if (fioCol >= 0 && scoreCol >= 0) {
            return new HeaderPos(fioCol, classCol, presenceCol, variantCol, scoreCol, 3);
        }
        return null;
    }

    private static boolean isMeaningfulResultRow(Row row, HeaderPos pos, Integer score) {
        String presence = pos.presenceCol >= 0 ? getCellText(row.getCell(pos.presenceCol)).trim().toLowerCase(Locale.ROOT) : "";
        boolean variantFilled = pos.variantCol >= 0 && !getCellText(row.getCell(pos.variantCol)).trim().isEmpty();
        boolean taskPointsFilled = hasTaskPoints(row, pos);

        if (score != null && score > 0) {
            return true;
        }

        // Явно отсутствовал на работе и других данных нет -> это пропуск, а не "0 баллов".
        if ((presence.contains("не был") || presence.contains("отсутств")) && !variantFilled && !taskPointsFilled) {
            return false;
        }

        // Во многих протоколах "Присутствие" может быть пустым:
        // если при этом нет варианта и нет баллов за задания, считаем, что ученик не писал.
        if (presence.isBlank() && !variantFilled && !taskPointsFilled && (score == null || score == 0)) {
            return false;
        }

        // Ключевое бизнес-правило: итог=0 считаем попыткой только если реально заполнялись баллы за задания.
        if (score != null && score == 0) {
            return taskPointsFilled;
        }

        // Для пустого итога допускаем загрузку только при наличии явных заполненных данных.
        return variantFilled || taskPointsFilled;
    }

    private static boolean hasTaskPoints(Row row, HeaderPos pos) {
        if (pos.scoreCol < 0) return false;
        int start = Math.max(0, Math.min(pos.fioCol, pos.scoreCol) + 1);
        int end = Math.max(pos.fioCol, pos.scoreCol) - 1;
        for (int c = start; c <= end; c++) {
            if (c == pos.classCol) continue;
            if (c == pos.presenceCol) continue;
            if (c == pos.variantCol) continue;
            String value = getCellText(row.getCell(c)).trim();
            if (!value.isEmpty()) {
                // "0" в промежуточных полях не считаем заполнением:
                // это часто шаблон/формула для не писавших работу.
                if (!"0".equals(value)) return true;
            }
        }
        return false;
    }

    private static String sortableClass(String className) {
        String normalized = normalizeClassName(className);
        return normalized.isBlank() ? "ZZZ" : normalized;
    }

    private static String extractSubject(Sheet infoSheet, Path file) {
        if (infoSheet == null) {
            LOG.warning("Лист 'Информация' отсутствует, предмет не определен: " + file.getFileName());
            return null;
        }

        // По согласованному правилу предмет всегда берется из B3.
        Row row = infoSheet.getRow(2);
        String b3 = row == null ? "" : getCellText(row.getCell(1));
        String subject = normalizeSubject(b3);
        if (subject != null) return subject;

        LOG.warning("Не удалось определить предмет из 'Информация'!B3: '" + b3 + "' (" + file.getFileName() + ")");
        return null;
    }

    private static Set<String> extractSubjects(String raw) {
        Set<String> out = new LinkedHashSet<>();
        String normalizedRaw = raw.replace(';', ',');

        if (normalizedRaw.toLowerCase(Locale.ROOT).contains("иностранный язык")) {
            // По согласованному правилу: всегда считаем иностранный язык как английский.
            out.add("Английский язык");
        }

        for (String subject : SUBJECT_ORDER) {
            if (normalizedRaw.toLowerCase(Locale.ROOT).contains(subject.toLowerCase(Locale.ROOT))) {
                out.add(subject);
            }
        }

        for (Map.Entry<String, String> alias : SUBJECT_ALIASES.entrySet()) {
            if (normalizedRaw.toLowerCase(Locale.ROOT).contains(alias.getKey().toLowerCase(Locale.ROOT))) {
                out.add(alias.getValue());
            }
        }

        return out;
    }

    private static String normalizeSubject(String value) {
        if (value == null) return null;
        String low = value.toLowerCase(Locale.ROOT);

        for (String subject : SUBJECT_ORDER) {
            if (low.contains(subject.toLowerCase(Locale.ROOT))) return subject;
        }
        for (Map.Entry<String, String> alias : SUBJECT_ALIASES.entrySet()) {
            if (low.contains(alias.getKey().toLowerCase(Locale.ROOT))) return alias.getValue();
        }
        return null;
    }

    private static String normalizeFio(String fio) {
        if (fio == null) return "";
        return fio.trim().replaceAll("\\s+", " ");
    }

    private static String normalizeClassName(String value) {
        if (value == null) return "";
        String normalized = value.trim()
                .replace('—', '-')
                .replace('–', '-')
                .replaceAll("\\s+", "")
                .toUpperCase(Locale.ROOT);
        // 9А -> 9-А
        normalized = normalized.replaceAll("^(\\d+)([А-ЯA-Z])$", "$1-$2");
        // 9 - А -> 9-А и аналогичные формы
        normalized = normalized.replaceAll("^(\\d+)\\-?([А-ЯA-Z])$", "$1-$2");
        return normalized;
    }

    private static String extractClassFromFileName(String fileName) {
        if (fileName == null) return "";
        String noExt = fileName.replaceAll("\\.xlsx$", "");
        var matcher = java.util.regex.Pattern.compile("(\\d+)\\s*[- ]?\\s*([А-ЯA-Z])").matcher(noExt.toUpperCase(Locale.ROOT));
        if (matcher.find()) {
            return matcher.group(1) + "-" + matcher.group(2);
        }
        return "";
    }

    private static Integer parseInt(String text) {
        if (text == null) return null;
        String normalized = text.trim().replace(',', '.');
        if (normalized.isEmpty()) return null;
        try {
            double d = Double.parseDouble(normalized);
            return (int) Math.round(d);
        } catch (Exception e) {
            return null;
        }
    }

    private static String getCellText(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> {
                double v = cell.getNumericCellValue();
                if (v == Math.floor(v)) yield String.valueOf((int) v);
                yield String.valueOf(v);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> {
                CellType cached = cell.getCachedFormulaResultType();
                if (cached == CellType.STRING) {
                    yield cell.getStringCellValue().trim();
                }
                if (cached == CellType.NUMERIC) {
                    double v = cell.getNumericCellValue();
                    if (v == Math.floor(v)) yield String.valueOf((int) v);
                    yield String.valueOf(v);
                }
                if (cached == CellType.BOOLEAN) {
                    yield String.valueOf(cell.getBooleanCellValue());
                }
                yield "";
            }
            case ERROR, BLANK, _NONE -> "";
            default -> "";
        };
    }

    private static CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private static CellStyle createFillStyle(Workbook workbook, short color) {
        CellStyle style = workbook.createCellStyle();
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(color);
        return style;
    }

    private static void autoSize(Sheet sheet, int cols) {
        for (int c = 0; c < cols; c++) {
            sheet.autoSizeColumn(c);
            int current = sheet.getColumnWidth(c);
            sheet.setColumnWidth(c, Math.min(current + 600, 12000));
        }
    }

    private record HeaderPos(int fioCol, int classCol, int presenceCol, int variantCol, int scoreCol, int dataStartRow) {
    }

    private static class ResultCell {
        final Integer score;
        final Integer grade;

        private ResultCell(Integer score, Integer grade) {
            this.score = score;
            this.grade = grade;
        }
    }

    private static class StudentRow {
        final String fio;
        String className = "";
        final Set<String> subjects = new LinkedHashSet<>();
        final Map<String, ResultCell> results = new HashMap<>();

        private StudentRow(String fio) {
            this.fio = fio;
        }
    }

    private record MissingItem(String subject, String fio) {
    }
}
