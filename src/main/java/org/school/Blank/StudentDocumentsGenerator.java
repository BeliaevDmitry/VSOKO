package org.school.Blank;

import com.itextpdf.text.Font;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;

import java.io.*;
import java.util.*;
import java.nio.file.*;
import java.util.List;

public class StudentDocumentsGenerator {

    private static final String EXCEL_PATH = "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ВСОКО\\Реестр контингента ОУ_19-12-2025 17-44.xlsx";
    private static final String OUTPUT_FOLDER = "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ВСОКО\\Готовые_бланки\\";

    // Классы, которые нужно обработать
    private static final Set<String> TARGET_CLASSES = Set.of("6");
    private static final String subject = "История";

    // Размеры и параметры
    private static final float PAGE_MARGIN = 25f;
    private static final float MARKER_SIZE = 10f;
    private static final float CELL_SIZE_MM = 5f; // 5 мм
    private static final float CELL_SIZE_PT = CELL_SIZE_MM * 2.83465f; // Конвертация мм в пункты (5 мм = 14.17 pt)
    private static final int CELLS_PER_ANSWER = 8;
    private static final int CORRECTION_ROWS = 6;
    private static final int VARIANT_CELLS = 2; // 2 клеточки для номера варианта

    // Расстояние между строками с ответами
    private static final float ROW_SPACING = 3f;

    // Полупрозрачный цвет для клеточек
    private static final BaseColor LIGHT_GRAY = new BaseColor(200, 200, 200);

    // Шрифт для использования в PDF
    private static BaseFont baseFont;
    private static Font regularFont;
    private static Font titleFont;
    private static Font smallFont;

    static {
        try {
            // Инициализация шрифтов один раз при загрузке класса
            baseFont = BaseFont.createFont("C:/Windows/Fonts/arial.ttf",
                    BaseFont.IDENTITY_H,
                    BaseFont.EMBEDDED);
            regularFont = new Font(baseFont, 12); // Шрифт 12 размера
            titleFont = new Font(baseFont, 16, Font.BOLD);
            smallFont = new Font(baseFont, 10);
        } catch (Exception e) {
            System.err.println("Ошибка инициализации шрифтов: " + e.getMessage());
            try {
                baseFont = BaseFont.createFont();
                regularFont = new Font(baseFont, 12);
                titleFont = new Font(baseFont, 16, Font.BOLD);
                smallFont = new Font(baseFont, 10);
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
    }

    public static void main(String[] args) {
        try {
            // Создаем папку для результатов
            Files.createDirectories(Paths.get(OUTPUT_FOLDER));

            // Читаем данные из Excel
            List<Student> students = readStudentsFromExcel();

            // Группируем студентов по классам
            Map<String, List<Student>> studentsByClass = groupStudentsByClass(students);

            // Обрабатываем каждый класс
            for (Map.Entry<String, List<Student>> entry : studentsByClass.entrySet()) {
                String className = entry.getKey();
                List<Student> classStudents = entry.getValue();

                System.out.println("Обработка класса " + className + ", количество учеников: " + classStudents.size());

                // Создаем PDF для класса
                createClassPDFs(className, classStudents);
            }

            System.out.println("Готово! Все PDF бланки созданы в папке: " + OUTPUT_FOLDER);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static List<Student> readStudentsFromExcel() throws IOException {
        List<Student> students = new ArrayList<>();

        try (FileInputStream file = new FileInputStream(EXCEL_PATH);
             Workbook workbook = new XSSFWorkbook(file)) {

            Sheet sheet = workbook.getSheetAt(0); // Первый лист

            // Пропускаем заголовок (предполагаем, что данные начинаются со второй строки)
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // Колонка C (индекс 2) - ФИО
                Cell fioCell = row.getCell(2);
                // Колонка P (индекс 15) - Класс
                Cell classCell = row.getCell(15);

                if (fioCell != null && classCell != null) {
                    String fio = getCellValueAsString(fioCell);
                    String className = getCellValueAsString(classCell).trim();

                    // Фильтруем только нужные классы
                    if (isTargetClass(className) && !fio.isEmpty()) {
                        students.add(new Student(fio, normalizeClassName(className)));
                    }
                }
            }
        }

        System.out.println("Прочитано " + students.size() + " учеников из Excel файла");
        return students;
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    // Для числовых значений классов
                    double value = cell.getNumericCellValue();
                    if (value == Math.floor(value)) {
                        return String.valueOf((int) value);
                    } else {
                        return String.valueOf(value);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    return String.valueOf(cell.getNumericCellValue());
                }
            default:
                return "";
        }
    }

    private static String normalizeClassName(String className) {
        // Убираем лишние пробелы и нормализуем
        return className.trim()
                .replaceAll("\\s+", " ")
                .replaceAll("\\s*-\\s*", "-")
                .replaceAll("[^\\dА-Яа-яA-Za-z\\s-]", "");
    }

    private static boolean isTargetClass(String className) {
        String normalized = normalizeClassName(className);

        // Ищем цифры в начале строки
        String firstDigits = normalized.replaceAll("^([\\d]+).*", "$1");

        // Проверяем, является ли первое число целевым классом
        return TARGET_CLASSES.contains(firstDigits);
    }

    private static Map<String, List<Student>> groupStudentsByClass(List<Student> students) {
        Map<String, List<Student>> grouped = new TreeMap<>();

        for (Student student : students) {
            String className = student.getClassName();
            grouped.computeIfAbsent(className, k -> new ArrayList<>()).add(student);
        }

        System.out.println("Найдено классов для обработки: " + grouped.keySet());
        return grouped;
    }

    private static void createClassPDFs(String className, List<Student> students) throws Exception {
        // Создаем папку для PDF файлов этого класса
        String classPdfFolder = OUTPUT_FOLDER + className + "_класс\\";
        Files.createDirectories(Paths.get(classPdfFolder));

        List<String> pdfFiles = new ArrayList<>();

        // Создаем PDF бланк для каждого студента
        for (int i = 0; i < students.size(); i++) {
            Student student = students.get(i);

            // Создаем имя файла без запрещенных символов
            String safeFio = student.getFio()
                    .replaceAll("[\\\\/:*?\"<>|]", "_")
                    .replaceAll("\\s+", "_");

            String pdfFileName = classPdfFolder +
                    String.format("%03d_", i + 1) +
                    safeFio + ".pdf";

            // Создаем PDF документ для студента
            createStudentPDF(student, pdfFileName);
            pdfFiles.add(pdfFileName);

            System.out.println("Создан PDF бланк для: " + student.getFio() +
                    " (" + student.getClassName() + ") - " +
                    (i + 1) + "/" + students.size());
        }

        System.out.println("Все PDF бланки для класса " + className + " сохранены в: " + classPdfFolder);

        // Объединяем PDF в один файл
        if (!pdfFiles.isEmpty()) {
            mergePDFs(pdfFiles, OUTPUT_FOLDER + "Все_бланки_" + className + "_класс.pdf");
        }

        System.out.println("=".repeat(60));
    }

    private static void createStudentPDF(Student student, String pdfFilePath) throws Exception {
        // Создаем PDF документ
        Document document = new Document(PageSize.A4, PAGE_MARGIN, PAGE_MARGIN, PAGE_MARGIN, PAGE_MARGIN);
        PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(pdfFilePath));

        document.open();

        try {
            // Добавляем 5 черных маркеров по углам
            addFiveCornerMarkers(document, writer);

            // Заголовок
            Paragraph title = new Paragraph("БЛАНК ОТВЕТОВ", titleFont);
            title.setAlignment(Element.ALIGN_CENTER);
            title.setSpacingAfter(10);
            document.add(title);

            document.add(Chunk.NEWLINE);

            // Информация об ученике
            Paragraph infoParagraph = new Paragraph();
            infoParagraph.add(new Chunk("Класс: ", regularFont));
            infoParagraph.add(new Chunk(student.getClassName(), regularFont));
            infoParagraph.add(new Chunk("    ФИО: ", regularFont));
            infoParagraph.add(new Chunk(student.getFio(), regularFont));
            infoParagraph.add(new Chunk("    Предмет: " + subject, regularFont));
            infoParagraph.setAlignment(Element.ALIGN_CENTER);
            document.add(infoParagraph);

            // Поле для номера варианта
            document.add(Chunk.NEWLINE);
            createVariantField(document);

            document.add(Chunk.NEWLINE);
            document.add(Chunk.NEWLINE);

            // Создаем таблицу для вопросов в 2 столбика
            createCompactQuestionsTable(document);

            document.add(Chunk.NEWLINE);

            // Зона для внесения исправлений
            createCompactCorrectionZone(document);

            document.add(Chunk.NEWLINE);

            // Инструкция
            Paragraph instruction = new Paragraph("Заполнять ЧЕРНОЙ ГЕЛЕВОЙ РУЧКОЙ. Каждый символ пишите в отдельной клетке.", smallFont);
            instruction.setAlignment(Element.ALIGN_CENTER);
            document.add(instruction);

        } finally {
            document.close();
            writer.close();
        }
    }

    private static void createVariantField(Document document) throws DocumentException {
        // Создаем параграф для поля "Вариант"
        Paragraph variantParagraph = new Paragraph();
        variantParagraph.add(new Chunk("Вариант: ", regularFont));
        variantParagraph.setAlignment(Element.ALIGN_CENTER);

        // Добавляем клеточки для номера варианта
        PdfPCell variantCell = createVariantCells();
        PdfPTable variantTable = new PdfPTable(1);
        variantTable.setWidthPercentage(10); // Узкая таблица для клеточек
        variantTable.setHorizontalAlignment(Element.ALIGN_CENTER);
        variantTable.addCell(variantCell);

        // Создаем общую таблицу для текста и клеточек
        PdfPTable containerTable = new PdfPTable(2);
        containerTable.setWidthPercentage(30);
        containerTable.setHorizontalAlignment(Element.ALIGN_CENTER);
        containerTable.setSpacingBefore(5f);

        // Текст "Вариант: "
        PdfPCell textCell = new PdfPCell(new Phrase("Вариант: ", regularFont));
        textCell.setBorder(Rectangle.NO_BORDER);
        textCell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        textCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        containerTable.addCell(textCell);

        // Клеточки для номера
        PdfPCell cellsCell = new PdfPCell();
        cellsCell.setBorder(Rectangle.NO_BORDER);
        cellsCell.addElement(variantTable);
        containerTable.addCell(cellsCell);

        document.add(containerTable);
    }

    private static PdfPCell createVariantCells() {
        PdfPCell cell = new PdfPCell();
        cell.setBorder(Rectangle.NO_BORDER);
        cell.setPadding(0);
        cell.setMinimumHeight(CELL_SIZE_PT + ROW_SPACING * 2);

        // Фиксированная ширина для 2 клеточек по 5мм = 10мм
        float totalWidth = CELL_SIZE_PT * VARIANT_CELLS + 1f;

        // Создаем таблицу с 2 клеточками для номера варианта
        PdfPTable cellsTable = new PdfPTable(VARIANT_CELLS);
        cellsTable.setTotalWidth(totalWidth);
        cellsTable.setLockedWidth(true);

        // Создаем 2 квадратные клеточки
        for (int i = 0; i < VARIANT_CELLS; i++) {
            PdfPCell variantCell = createSingleSquareCell();
            cellsTable.addCell(variantCell);
        }

        cell.addElement(cellsTable);
        return cell;
    }

    private static void addFiveCornerMarkers(Document document, PdfWriter writer) {
        PdfContentByte canvas = writer.getDirectContent();
        float pageWidth = document.getPageSize().getWidth();
        float pageHeight = document.getPageSize().getHeight();

        // Заливаем маркеры черным цветом
        canvas.setColorFill(BaseColor.BLACK);

        // 1. Левый верхний угол
        canvas.rectangle(PAGE_MARGIN - 5, pageHeight - PAGE_MARGIN, MARKER_SIZE, MARKER_SIZE);
        canvas.fill();

        // 2. Правый верхний угол
        canvas.rectangle(pageWidth - PAGE_MARGIN - MARKER_SIZE + 5, pageHeight - PAGE_MARGIN, MARKER_SIZE, MARKER_SIZE);
        canvas.fill();

        // 3. Левый нижний угол
        canvas.rectangle(PAGE_MARGIN - 5, PAGE_MARGIN, MARKER_SIZE, MARKER_SIZE);
        canvas.fill();

        // 4. Правый нижний угол
        canvas.rectangle(pageWidth - PAGE_MARGIN - MARKER_SIZE + 5, PAGE_MARGIN, MARKER_SIZE, MARKER_SIZE);
        canvas.fill();

        // 5. Центр верхнего края (для определения ориентации)
        canvas.rectangle(pageWidth / 2 - MARKER_SIZE / 2, pageHeight - PAGE_MARGIN, MARKER_SIZE, MARKER_SIZE);
        canvas.fill();
    }

    private static void createCompactQuestionsTable(Document document) throws DocumentException {
        // Создаем основную таблицу с 2 колонками
        PdfPTable mainTable = new PdfPTable(2);
        mainTable.setWidthPercentage(70);
        mainTable.setSpacingBefore(5f);
        mainTable.setSpacingAfter(5f);
        mainTable.setKeepTogether(true);

        // Ширина колонок: 50%, 50%
        float[] columnWidths = {50f, 50f};
        mainTable.setWidths(columnWidths);

        // Создаем левый столбец (вопросы 1-15)
        PdfPCell leftColumn = createCompactQuestionsColumn(1, 15);
        mainTable.addCell(leftColumn);

        // Создаем правый столбец (вопросы 16-30)
        PdfPCell rightColumn = createCompactQuestionsColumn(16, 30);
        mainTable.addCell(rightColumn);

        document.add(mainTable);
    }

    private static PdfPCell createCompactQuestionsColumn(int startQuestion, int endQuestion) throws DocumentException {
        PdfPCell cell = new PdfPCell();
        cell.setBorder(Rectangle.NO_BORDER);
        cell.setPadding(0);
        cell.setMinimumHeight(0);

        // Создаем внутреннюю таблицу для вопросов
        PdfPTable innerTable = new PdfPTable(2);
        innerTable.setWidthPercentage(100);
        innerTable.setSpacingBefore(0);
        innerTable.setSpacingAfter(0);

        // Ширина колонок: 12% для номера, 88% для клеточек
        float[] innerWidths = {12f, 88f};
        innerTable.setWidths(innerWidths);

        // Добавляем вопросы в столбец
        for (int questionNum = startQuestion; questionNum <= endQuestion; questionNum++) {
            // Ячейка с номером вопроса (маленький шрифт)
            PdfPCell numberCell = new PdfPCell(new Phrase(String.valueOf(questionNum) + ".", smallFont));
            numberCell.setBorder(Rectangle.NO_BORDER);
            numberCell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            numberCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            numberCell.setPaddingRight(1f);
            numberCell.setPaddingTop(ROW_SPACING);
            numberCell.setPaddingBottom(ROW_SPACING);
            numberCell.setMinimumHeight(CELL_SIZE_PT + ROW_SPACING * 2);

            // Ячейка с 8 клеточками для ответа
            PdfPCell answerCell = createFixedWidthAnswerCells();
            answerCell.setPaddingTop(ROW_SPACING);
            answerCell.setPaddingBottom(ROW_SPACING);

            innerTable.addCell(numberCell);
            innerTable.addCell(answerCell);
        }

        cell.addElement(innerTable);
        return cell;
    }

    private static PdfPCell createFixedWidthAnswerCells() {
        PdfPCell cell = new PdfPCell();
        cell.setBorder(Rectangle.NO_BORDER);
        cell.setPadding(0);
        cell.setMinimumHeight(CELL_SIZE_PT + ROW_SPACING * 2);

        // Фиксированная ширина для 8 клеточек по 5мм = 40мм + отступы
        float totalWidth = CELL_SIZE_PT * CELLS_PER_ANSWER + 2f; // 8 * 14.17pt + отступы

        // Создаем таблицу с 8 клеточками
        PdfPTable cellsTable = new PdfPTable(CELLS_PER_ANSWER);
        cellsTable.setTotalWidth(totalWidth);
        cellsTable.setLockedWidth(true); // Фиксируем ширину

        // Устанавливаем фиксированную ширину для каждой колонки
        float[] widths = new float[CELLS_PER_ANSWER];
        for (int i = 0; i < CELLS_PER_ANSWER; i++) {
            widths[i] = CELL_SIZE_PT;
        }
        try {
            cellsTable.setWidths(widths);
        } catch (DocumentException e) {
            throw new RuntimeException(e);
        }

        // Создаем 8 квадратных клеточек
        for (int i = 0; i < CELLS_PER_ANSWER; i++) {
            PdfPCell singleCell = createSingleSquareCell();
            cellsTable.addCell(singleCell);
        }

        cell.addElement(cellsTable);
        return cell;
    }

    private static PdfPCell createSingleSquareCell() {
        // Создаем квадратную клеточку 5x5 мм
        PdfPCell cell = new PdfPCell();
        cell.setFixedHeight(CELL_SIZE_PT);
        cell.setBorder(Rectangle.BOX);
        cell.setBorderWidth(0.5f);
        cell.setBorderColor(LIGHT_GRAY);
        cell.setBackgroundColor(new BaseColor(255, 255, 255, 220)); // Полупрозрачный
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);

        // Добавляем направляющую линию
        cell.setPhrase(new Phrase(" ", new Font(baseFont, 6)));

        return cell;
    }

    private static void createCompactCorrectionZone(Document document) throws DocumentException {
        // Создаем зону для внесения исправлений
        Paragraph correctionTitle = new Paragraph("Замена ошибочных ответов:", smallFont);
        correctionTitle.setSpacingBefore(5f);
        correctionTitle.setAlignment(Element.ALIGN_CENTER);
        document.add(correctionTitle);

        // Создаем компактную таблицу для исправлений
        PdfPTable correctionTable = new PdfPTable(2);
        correctionTable.setWidthPercentage(30);
        correctionTable.setSpacingBefore(3f);
        correctionTable.setSpacingAfter(3f);

        // Ширина колонок: 20% для номера, 80% для ответа
        float[] widths = {20f, 80f};
        correctionTable.setWidths(widths);

        // Добавляем 6 компактных строк для исправлений
        for (int i = 1; i <= CORRECTION_ROWS; i++) {
            // Номер задания (2 клеточки)
            PdfPCell taskNumberCell = createCompactTaskNumberCells();
            taskNumberCell.setPaddingTop(ROW_SPACING);
            taskNumberCell.setPaddingBottom(ROW_SPACING);

            // Новый ответ (8 клеточек)
            PdfPCell newAnswerCell = createFixedWidthAnswerCells();
            newAnswerCell.setPaddingTop(ROW_SPACING);
            newAnswerCell.setPaddingBottom(ROW_SPACING);

            correctionTable.addCell(taskNumberCell);
            correctionTable.addCell(newAnswerCell);
        }

        document.add(correctionTable);

        // Пояснение
        Paragraph note = new Paragraph("Если нужно исправить ответ, впишите номер задания и новый ответ в клеточки справа",
                new Font(baseFont, 8));
        note.setAlignment(Element.ALIGN_CENTER);
        note.setSpacingBefore(2f);
        document.add(note);
    }

    private static PdfPCell createCompactTaskNumberCells() {
        PdfPCell cell = new PdfPCell();
        cell.setBorder(Rectangle.NO_BORDER);
        cell.setPadding(0);
        cell.setMinimumHeight(CELL_SIZE_PT + ROW_SPACING * 2);

        // Фиксированная ширина для 2 клеточек по 5мм = 10мм
        float totalWidth = CELL_SIZE_PT * 2 + 1f;

        // Создаем таблицу с 2 клеточками
        PdfPTable cellsTable = new PdfPTable(2);
        cellsTable.setTotalWidth(totalWidth);
        cellsTable.setLockedWidth(true);

        // Создаем 2 квадратные клеточки
        for (int i = 0; i < 2; i++) {
            PdfPCell numberCell = createSingleSquareCell();
            cellsTable.addCell(numberCell);
        }

        cell.addElement(cellsTable);
        return cell;
    }

    private static void mergePDFs(List<String> pdfFiles, String outputPath) throws Exception {
        if (pdfFiles.isEmpty()) {
            System.out.println("Нет PDF файлов для объединения");
            return;
        }

        System.out.println("Объединение " + pdfFiles.size() + " PDF файлов в один...");

        Document document = new Document();
        PdfCopy copy = new PdfCopy(document, new FileOutputStream(outputPath));
        document.open();

        try {
            for (String pdfFile : pdfFiles) {
                PdfReader reader = new PdfReader(pdfFile);
                int n = reader.getNumberOfPages();

                for (int page = 0; page < n; page++) {
                    copy.addPage(copy.getImportedPage(reader, page + 1));
                }

                copy.freeReader(reader);
                reader.close();
            }
        } catch (Exception e) {
            System.err.println("Ошибка при объединении PDF: " + e.getMessage());
            throw e;
        } finally {
            document.close();
        }

        System.out.println("Объединенный PDF создан: " + outputPath);
    }

    // Вспомогательный класс для хранения данных студента
    static class Student {
        private final String fio;
        private final String className;

        public Student(String fio, String className) {
            this.fio = fio;
            this.className = className;
        }

        public String getFio() {
            return fio;
        }

        public String getClassName() {
            return className;
        }

        @Override
        public String toString() {
            return fio + " (" + className + ")";
        }
    }
}