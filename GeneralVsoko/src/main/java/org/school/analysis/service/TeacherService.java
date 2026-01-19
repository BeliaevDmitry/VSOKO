package org.school.analysis.service;

import jakarta.persistence.EntityManager;
import jakarta.persistence.PersistenceContext;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.text.similarity.LevenshteinDistance;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.school.analysis.config.AppConfig;
import org.school.analysis.model.Teacher;
import org.school.analysis.repository.TeacherRepository;
import org.school.analysis.util.TeacherMatcher;
import org.school.analysis.util.TeacherNameNormalizer;
import org.springframework.dao.DataAccessException;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import jakarta.annotation.PostConstruct;

import java.io.File;
import java.io.FileInputStream;
import java.time.LocalDateTime;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.stream.Collectors;

@Service
@Slf4j
@RequiredArgsConstructor
public class TeacherService {

    private final TeacherRepository teacherRepository;
    private static final LevenshteinDistance LEVENSHTEIN = new LevenshteinDistance();
    private final TeacherMatcher teacherMatcher; // Добавляем
    private final Map<String, Teacher> teachersCache = new ConcurrentHashMap<>();
    private final Map<String, String> normalizedToFullName = new ConcurrentHashMap<>();
    private final Map<String, List<String>> lastNameVariants = new ConcurrentHashMap<>();

    private String lastFileHash = "";
    private LocalDateTime lastImportTime;


    @PersistenceContext
    private EntityManager entityManager;

    /**
     * Инициализация кэша при запуске
     */
    public void init() {
        loadTeachersFromDatabase();
    }

    /**
     * Расчет хэша файла для отслеживания изменений
     */
    private String calculateFileHash(File file) {
        try {
            return String.valueOf(file.lastModified() + file.length());
        } catch (Exception e) {
            log.warn("Не удалось рассчитать хэш файла: {}", e.getMessage());
            return "";
        }
    }

    /**
     * Импорт учителей из Excel файла
     */
    @Transactional(propagation = Propagation.REQUIRES_NEW, noRollbackFor = {Exception.class})
    public void importTeachersFromExcel(String school) {
        String filePath = AppConfig.INPUT_TEACHER_NAME.replace("{школа}", school);
        File teacherFile = new File(filePath);

        if (!teacherFile.exists()) {
            log.warn("Файл учителей не найден: {}. Пропускаем импорт.", filePath);
            return;
        }

        log.info("Начинаем импорт учителей из файла: {}", filePath);

        try (FileInputStream file = new FileInputStream(teacherFile);
             Workbook workbook = new XSSFWorkbook(file)) {

            Sheet sheet = workbook.getSheetAt(0);
            int importedCount = 0;
            int skippedCount = 0;
            int errorCount = 0;

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }

                Cell nameCell = row.getCell(1);
                if (nameCell == null || isCellEmpty(nameCell)) {
                    continue;
                }

                String fullName = getCellValue(nameCell).trim();
                if (fullName.isEmpty()) {
                    continue;
                }

                try {
                    Teacher teacher = createOrUpdateTeacherInternal(fullName, filePath);
                    if (teacher != null) {
                        importedCount++;
                    } else {
                        skippedCount++;
                    }
                } catch (Exception e) {
                    log.error("Ошибка при импорте учителя '{}': {}", fullName, e.getMessage());
                    errorCount++;
                }
            }

            loadTeachersFromDatabase();
            lastFileHash = calculateFileHash(teacherFile);
            lastImportTime = LocalDateTime.now();

            log.info("Импорт учителей завершен. Импортировано: {}, Пропущено: {}, Ошибок: {}",
                    importedCount, skippedCount, errorCount);

        } catch (Exception e) {
            log.error("Ошибка при импорте учителей из файла: {}", e.getMessage(), e);
        }
    }

    /**
     * Создает или обновляет учителя в БД (для внешних вызовов)
     */
    @Transactional
    public Teacher createOrUpdateTeacher(String fullName, String sourceFile) {
        return createOrUpdateTeacherInternal(fullName, sourceFile);
    }

    /**
     * Внутренний метод создания/обновления учителя
     */
    private Teacher createOrUpdateTeacherInternal(String fullName, String sourceFile) {
        try {
            String normalized = TeacherNameNormalizer.normalize(fullName);
            Optional<Teacher> existingTeacher = teacherRepository.findByNormalizedFullName(normalized);

            if (existingTeacher.isPresent()) {
                Teacher teacher = existingTeacher.get();
                if (!teacher.getIsActive()) {
                    teacher.setIsActive(true);
                    teacher.setSourceFile(sourceFile);
                    teacher.setUpdatedAt(LocalDateTime.now());
                    teacher = teacherRepository.save(teacher);
                    log.debug("Активирован существующий учитель: {}", fullName);
                }
                return teacher;
            }

            String[] parts = normalized.split(" ");
            if (parts.length < 2) {
                log.warn("Некорректное ФИО для импорта: {}", fullName);
                return null;
            }

            Teacher teacher = Teacher.builder()
                    .fullName(fullName)
                    .normalizedFullName(normalized)
                    .lastName(parts[0])
                    .firstName(parts.length > 1 ? parts[1] : "")
                    .middleName(parts.length > 2 ? parts[2] : "")
                    .shortName(TeacherNameNormalizer.getShortName(fullName))
                    .isActive(true)
                    .sourceFile(sourceFile)
                    .createdAt(LocalDateTime.now())
                    .build();

            Teacher saved = teacherRepository.save(teacher);
            log.debug("Создан новый учитель: {}", fullName);
            return saved;
        } catch (Exception e) {
            log.error("Ошибка при создании/обновлении учителя '{}': {}", fullName, e.getMessage());
            return null;
        }
    }

    /**
     * Обновление времени последнего использования учителя
     */
    @Transactional(propagation = Propagation.REQUIRES_NEW)
    public void updateTeacherLastSeen(String teacherName) {
        try {
            String normalized = TeacherNameNormalizer.normalize(teacherName);
            Optional<Teacher> teacherOpt = teacherRepository.findByNormalizedFullName(normalized);

            if (teacherOpt.isEmpty()) {
                teacherOpt = findFuzzyMatchInDatabase(teacherName);
            }

            if (teacherOpt.isPresent()) {
                Teacher teacher = teacherOpt.get();
                teacher.setLastSeenInReport(LocalDateTime.now());
                teacherRepository.save(teacher);
                log.debug("Обновлен lastSeen для учителя: {}", teacher.getFullName());
            }
        } catch (Exception e) {
            log.warn("Не удалось обновить lastSeen для учителя {}: {}", teacherName, e.getMessage());
        }
    }

    /**
     * Обновление lastSeen для первого учителя с данной фамилией
     */
    @Transactional(propagation = Propagation.REQUIRES_NEW)
    public void updateLastSeenForLastName(String lastName) {
        try {
            List<Teacher> teachers = teacherRepository.findByLastNameIgnoreCase(lastName);
            if (!teachers.isEmpty()) {
                Teacher teacher = teachers.get(0);
                teacher.setLastSeenInReport(LocalDateTime.now());
                teacherRepository.save(teacher);
                log.debug("Обновлен lastSeen для учителя по фамилии {}: {}", lastName, teacher.getFullName());
            }
        } catch (Exception e) {
            log.warn("Не удалось обновить lastSeen по фамилии {}: {}", lastName, e.getMessage());
        }
    }


    /**
     * Новый метод: получение полного имени учителя с использованием TeacherMatcher
     */
    public Optional<String> getFullTeacherName(String teacherNameFromReport) {
        if (teacherNameFromReport == null || teacherNameFromReport.trim().isEmpty()) {
            return Optional.empty();
        }

        // Пробуем старый метод для обратной совместимости
        Optional<String> oldResult = getFullTeacherNameOld(teacherNameFromReport);
        if (oldResult.isPresent()) {
            return oldResult;
        }

        // Используем новый matcher
        List<Teacher> allTeachers = teacherRepository.findAll();
        Optional<Teacher> matchedTeacher = teacherMatcher.findMatchingTeacher(teacherNameFromReport, allTeachers);

        return matchedTeacher.map(Teacher::getFullName);
    }

    /**
     * Старый метод (для обратной совместимости)
     */
    private Optional<String> getFullTeacherNameOld(String teacherNameFromReport) {
        if (teacherNameFromReport == null) {
            return Optional.empty();
        }

        String normalized = TeacherNameNormalizer.normalize(teacherNameFromReport);
        if (normalizedToFullName.containsKey(normalized)) {
            return Optional.of(normalizedToFullName.get(normalized));
        }

        Optional<Teacher> teacherOpt = findFuzzyMatchInDatabase(teacherNameFromReport);
        return teacherOpt.map(Teacher::getFullName);
    }

    /**
     * Обновленный метод проверки валидности учителя
     */
    public boolean isTeacherValid(String teacherNameFromReport) {
        if (teacherNameFromReport == null || teacherNameFromReport.trim().isEmpty()) {
            log.warn("Имя учителя пустое");
            return false;
        }

        // Используем новый matcher
        List<Teacher> allTeachers = teacherRepository.findAll();
        Optional<Teacher> matchedTeacher = teacherMatcher.findMatchingTeacher(teacherNameFromReport, allTeachers);

        if (matchedTeacher.isPresent()) {
            Teacher teacher = matchedTeacher.get();
            log.debug("✓ Найден учитель для '{}': '{}'",
                    teacherNameFromReport, teacher.getFullName());
            updateTeacherLastSeen(teacherNameFromReport);
            return true;
        }

        log.warn("✗ Учитель '{}' не найден", teacherNameFromReport);
        return false;
    }
    /**
     * Загрузка учителей из БД в кэш
     */
    private void loadTeachersFromDatabase() {
        log.info("Загрузка учителей из базы данных...");

        teachersCache.clear();
        normalizedToFullName.clear();
        lastNameVariants.clear();

        List<Teacher> activeTeachers = teacherRepository.findByIsActiveTrue();

        // ДОПОЛНИТЕЛЬНАЯ ПРОВЕРКА
        List<Teacher> validTeachers = activeTeachers.stream()
                .filter(t -> t.getId() != null)
                .peek(t -> {
                    if (t.getNormalizedFullName() == null) {
                        log.warn("Учитель ID {} имеет null normalizedFullName: {}",
                                t.getId(), t.getFullName());
                    }
                })
                .collect(Collectors.toList());

        if (validTeachers.size() < activeTeachers.size()) {
            log.warn("Найдено {} учителей с null ID из {}",
                    activeTeachers.size() - validTeachers.size(), activeTeachers.size());
        }

        for (Teacher teacher : validTeachers) {
            teachersCache.put(teacher.getFullName(), teacher);
            normalizedToFullName.put(teacher.getNormalizedFullName(), teacher.getFullName());

            if (teacher.getLastName() != null && !teacher.getLastName().isEmpty()) {
                String lastNameLower = teacher.getLastName().toLowerCase();
                lastNameVariants
                        .computeIfAbsent(lastNameLower, k -> new ArrayList<>())
                        .add(teacher.getFullName());
            }
        }

        log.info("Загружено {} активных учителей из БД (валидных: {})",
                activeTeachers.size(), validTeachers.size());
    }

    /**
     * Обновление кэша учителей
     */
    @Scheduled(fixedDelay = 300000)
    public void refreshCache() {
        log.debug("Обновление кэша учителей...");
        loadTeachersFromDatabase();
    }


    /**
     * Расширенный поиск с учетом возможных вариантов написания
     */
    private Optional<Teacher> findExtendedMatch(String teacherName) {
        String normalized = TeacherNameNormalizer.normalize(teacherName);

        // Ищем все учителей с похожей фамилией
        String[] parts = normalized.split(" ");
        if (parts.length == 0) {
            return Optional.empty();
        }

        // 1. Ищем по фамилии (нечеткое совпадение)
        List<Teacher> allTeachers = teacherRepository.findAll();
        List<Teacher> potentialMatches = new ArrayList<>();

        for (Teacher teacher : allTeachers) {
            String teacherNormalized = teacher.getNormalizedFullName();

            // Проверяем совпадение фамилий
            if (teacherNormalized.contains(parts[0]) || parts[0].contains(teacherNormalized.split(" ")[0])) {
                potentialMatches.add(teacher);
            }
        }

        // 2. Ищем по всем словам
        if (potentialMatches.isEmpty()) {
            for (Teacher teacher : allTeachers) {
                String teacherNormalized = teacher.getNormalizedFullName();
                for (String part : parts) {
                    if (teacherNormalized.contains(part) && part.length() > 2) {
                        potentialMatches.add(teacher);
                        break;
                    }
                }
            }
        }

        // 3. Выбираем лучшего кандидата
        if (!potentialMatches.isEmpty()) {
            // Для имени "Шпота Юлия Андреевна" ищем совпадение по имени и отчеству
            for (Teacher teacher : potentialMatches) {
                String teacherNormalized = teacher.getNormalizedFullName();
                String[] teacherParts = teacherNormalized.split(" ");

                if (teacherParts.length >= 3 && parts.length >= 3) {
                    // Проверяем совпадение имени и отчества
                    boolean nameMatches = teacherParts[1].equals(parts[1]) ||
                            (teacherParts[1].startsWith(parts[1]) && parts[1].length() > 1);
                    boolean middleNameMatches = teacherParts[2].equals(parts[2]) ||
                            (teacherParts[2].startsWith(parts[2]) && parts[2].length() > 1);

                    if (nameMatches && middleNameMatches) {
                        return Optional.of(teacher);
                    }
                }
            }

            // Возвращаем первого подходящего
            return Optional.of(potentialMatches.get(0));
        }

        return Optional.empty();
    }

    /**
     * Нечеткий поиск в БД
     */
    private Optional<Teacher> findFuzzyMatchInDatabase(String teacherName) {
        String normalized = TeacherNameNormalizer.normalize(teacherName);
        String[] parts = normalized.split(" ");

        if (parts.length == 0) {
            return Optional.empty();
        }

        log.debug("Нечеткий поиск для '{}' (части: {})", teacherName, Arrays.toString(parts));

        // Ищем по фамилии в БД
        List<Teacher> sameLastNameTeachers = teacherRepository.findByLastNameIgnoreCase(parts[0]);

        log.debug("Найдено учителей с фамилией '{}': {}", parts[0], sameLastNameTeachers.size());

        if (sameLastNameTeachers.isEmpty()) {
            // Попробуем найти по любому из слов
            for (String part : parts) {
                if (part.length() > 2) { // Ищем только значимые слова
                    List<Teacher> teachersByWord = teacherRepository.findByLastNameContainingIgnoreCase(part);
                    if (!teachersByWord.isEmpty()) {
                        sameLastNameTeachers = teachersByWord;
                        log.debug("Найдено по слову '{}': {}", part, teachersByWord.size());
                        break;
                    }
                }
            }
        }

        if (sameLastNameTeachers.isEmpty()) {
            return Optional.empty();
        }

        // Если только один учитель с такой фамилией, возвращаем его
        if (sameLastNameTeachers.size() == 1) {
            Teacher teacher = sameLastNameTeachers.get(0);
            log.debug("Один учитель с фамилией '{}': {}", parts[0], teacher.getFullName());
            return Optional.of(teacher);
        }

        log.debug("Несколько учителей с фамилией '{}': {}", parts[0],
                sameLastNameTeachers.stream()
                        .map(Teacher::getFullName)
                        .collect(Collectors.toList()));

        // Ищем лучшего кандидата по инициалам
        Teacher bestMatch = findBestMatchByInitials(normalized, parts, sameLastNameTeachers);

        if (bestMatch != null) {
            log.debug("Найден лучший кандидат по инициалам: {}", bestMatch.getFullName());
            return Optional.of(bestMatch);
        }

        // Если не нашли по инициалам, проверяем схожесть через Levenshtein
        return findBestMatchByLevenshtein(normalized, sameLastNameTeachers);
    }

    /**
     * Поиск лучшего совпадения по инициалам
     */
    private Teacher findBestMatchByInitials(String normalized, String[] parts, List<Teacher> candidates) {
        if (parts.length <= 1) {
            return null;
        }

        Teacher bestMatch = null;
        int bestScore = -1;

        for (Teacher teacher : candidates) {
            String teacherNormalized = teacher.getNormalizedFullName();
            String[] teacherParts = teacherNormalized.split(" ");

            if (teacherParts.length <= 1) {
                continue;
            }

            int score = calculateInitialsScore(parts, teacherParts);
            if (score > bestScore) {
                bestScore = score;
                bestMatch = teacher;
            }
        }

        return bestScore > 0 ? bestMatch : null;
    }

    /**
     * Расчет оценки совпадения инициалов
     */
    private int calculateInitialsScore(String[] searchParts, String[] teacherParts) {
        int score = 0;

        for (int i = 1; i < Math.min(searchParts.length, teacherParts.length); i++) {
            if (searchParts[i].length() > 0 && teacherParts[i].length() > 0) {
                if (searchParts[i].charAt(0) == teacherParts[i].charAt(0)) {
                    score += 2;
                } else if (i == 1 && searchParts[i].length() > 1 && teacherParts[i].length() > 1) {
                    if (teacherParts[i].startsWith(searchParts[i])) {
                        score += 3;
                    }
                }
            }
        }

        return score;
    }

    /**
     * Поиск лучшего совпадения по расстоянию Левенштейна
     */
    private Optional<Teacher> findBestMatchByLevenshtein(String normalized, List<Teacher> candidates) {
        Teacher bestMatch = null;
        double bestSimilarity = 0.0;

        for (Teacher teacher : candidates) {
            String teacherNormalized = teacher.getNormalizedFullName();
            Integer distance = LEVENSHTEIN.apply(normalized, teacherNormalized);
            int maxLength = Math.max(normalized.length(), teacherNormalized.length());

            if (maxLength > 0) {
                double similarity = 1.0 - (double) distance / maxLength;
                if (similarity > bestSimilarity) {
                    bestSimilarity = similarity;
                    bestMatch = teacher;
                }
            }
        }

        if (bestSimilarity >= 0.8 && bestMatch != null) {
            log.debug("Найдено совпадение по Левенштейну: {} (схожесть: {})",
                    bestMatch.getFullName(), bestSimilarity);
            return Optional.of(bestMatch);
        }

        return Optional.empty();
    }


    /**
     * Получение всех активных учителей
     */
    public List<Teacher> getAllActiveTeachers() {
        return new ArrayList<>(teachersCache.values());
    }

    /**
     * Добавление нового учителя вручную
     */
    @Transactional
    public Teacher addTeacherManually(String fullName) {
        Teacher teacher = createOrUpdateTeacherInternal(fullName, "manual_add");
        if (teacher != null) {
            loadTeachersFromDatabase();
        }
        return teacher;
    }

    /**
     * Деактивация учителя
     */
    @Transactional
    public void deactivateTeacher(Long teacherId) {
        try {
            teacherRepository.findById(teacherId).ifPresent(teacher -> {
                teacher.setIsActive(false);
                teacherRepository.save(teacher);
                loadTeachersFromDatabase();
            });
        } catch (DataAccessException e) {
            log.error("Ошибка базы данных при деактивации учителя {}: {}", teacherId, e.getMessage());
            throw e;
        } catch (Exception e) {
            log.error("Ошибка при деактивации учителя {}: {}", teacherId, e.getMessage());
        }
    }

    /**
     * Получение значения ячейки
     */
    private String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue().trim();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue().toString();
            } else {
                return String.valueOf((int) cell.getNumericCellValue());
            }
        } else {
            return "";
        }
    }

    /**
     * Проверка пустая ли ячейка
     */
    private boolean isCellEmpty(Cell cell) {
        if (cell == null) {
            return true;
        }

        if (cell.getCellType() == CellType.BLANK) {
            return true;
        }

        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue().trim().isEmpty();
        }

        return false;
    }

    /**
     * Получить статистику по учителям
     */
    public Map<String, Object> getStatistics() {
        Map<String, Object> stats = new HashMap<>();
        stats.put("totalInDb", teacherRepository.count());
        stats.put("activeTeachers", teacherRepository.countActiveTeachers());
        stats.put("inactiveTeachers", teacherRepository.count() - teacherRepository.countActiveTeachers());
        stats.put("cachedTeachers", teachersCache.size());
        stats.put("lastImportTime", lastImportTime);
        stats.put("lastFileHash", lastFileHash);
        stats.put("lastNameVariants", lastNameVariants.size());
        return stats;
    }
}