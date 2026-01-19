package org.school.analysis.util;

import lombok.extern.slf4j.Slf4j;
import org.school.analysis.model.Teacher;
import org.springframework.stereotype.Component;

import java.util.*;
import java.util.regex.Pattern;

@Slf4j
@Component
public class TeacherMatcher {

    private static final Pattern MULTIPLE_SPACES = Pattern.compile("\\s+");
    private static final Pattern PUNCTUATION = Pattern.compile("[.,!?;:()\\[\\]{}<>\\/\\\\\"'`~@#$%^&*_+=|]");

    /**
     * Простая нормализация для сравнения (убираем всё лишнее)
     */
    public String normalizeForComparison(String name) {
        if (name == null) {
            return "";
        }

        String result = name.toLowerCase();

        // 1. Удаляем пунктуацию
        result = PUNCTUATION.matcher(result).replaceAll("");

        // 2. Заменяем все последовательности пробелов одним пробелом
        result = MULTIPLE_SPACES.matcher(result.trim()).replaceAll(" ");

        // 3. Убираем лишние пробелы
        result = result.trim();

        // 4. Заменяем латинские буквы на русские
        result = replaceLatinWithRussian(result);

        return result;
    }

    /**
     * Замена латинских букв на русские
     */
    private String replaceLatinWithRussian(String text) {
        char[] chars = text.toCharArray();
        for (int i = 0; i < chars.length; i++) {
            switch (chars[i]) {
                case 'a': chars[i] = 'а'; break;
                case 'b': chars[i] = 'б'; break;
                case 'c': chars[i] = 'с'; break;
                case 'e': chars[i] = 'е'; break;
                case 'h': chars[i] = 'н'; break;
                case 'k': chars[i] = 'к'; break;
                case 'm': chars[i] = 'м'; break;
                case 'o': chars[i] = 'о'; break;
                case 'p': chars[i] = 'р'; break;
                case 't': chars[i] = 'т'; break;
                case 'x': chars[i] = 'х'; break;
                case 'y': chars[i] = 'у'; break;
            }
        }
        return new String(chars);
    }

    /**
     * Основной метод: проверяем все возможные варианты сопоставления
     * @param teacherNameFromReport - ФИО из отчета
     * @param teacherFromDatabase - учитель из базы
     * @return true если найдено совпадение
     */
    public boolean matchTeacher(String teacherNameFromReport, Teacher teacherFromDatabase) {
        if (teacherNameFromReport == null || teacherFromDatabase == null) {
            return false;
        }

        String reportNormalized = normalizeForComparison(teacherNameFromReport);
        String dbFullName = teacherFromDatabase.getFullName();
        String dbNormalized = normalizeForComparison(dbFullName);

        log.debug("Сопоставление: '{}' (norm: '{}') с '{}' (norm: '{}')",
                teacherNameFromReport, reportNormalized, dbFullName, dbNormalized);

        // 1. Простое сравнение (после нормализации)
        if (reportNormalized.equals(dbNormalized)) {
            log.debug("✓ Прямое совпадение после нормализации");
            return true;
        }

        // 2. Разбиваем на части
        String[] reportParts = reportNormalized.split(" ");
        String[] dbParts = dbNormalized.split(" ");

        if (reportParts.length == 0 || dbParts.length == 0) {
            return false;
        }

        // 3. Проверяем совпадение фамилии
        String reportSurname = reportParts[0];
        String dbSurname = dbParts[0];

        if (!reportSurname.equals(dbSurname)) {
            // Фамилии не совпадают - дальше не проверяем
            return false;
        }

        // 4. Фамилии совпадают, проверяем остальное

        // Случай 1: В отчете только фамилия
        if (reportParts.length == 1) {
            log.debug("✓ Совпадение по фамилии: '{}'", reportSurname);
            return true;
        }

        // Случай 2: В отчете "Фамилия И.О."
        if (isShortFormat(teacherNameFromReport)) {
            return matchShortFormat(teacherNameFromReport, teacherFromDatabase);
        }

        // Случай 3: Разный порядок слов (например, "И И Иванов")
        if (reportParts.length >= 2) {
            return matchDifferentOrder(reportNormalized, dbNormalized, reportParts, dbParts);
        }

        return false;
    }

    /**
     * Проверяем формат "Фамилия И.О."
     */
    private boolean isShortFormat(String name) {
        // Паттерн: Фамилия + одна или две буквы с точками
        return name.matches("^[А-Яа-яёЁ]+\\s+[А-Яа-яёЁ]\\.(\\s*[А-Яа-яёЁ]\\.)?$");
    }

    /**
     * Сопоставление для формата "Фамилия И.О."
     */
    private boolean matchShortFormat(String shortName, Teacher teacher) {
        try {
            // Извлекаем инициалы из короткого имени
            String initials = extractInitialsFromShortName(shortName);
            if (initials.isEmpty()) {
                return false;
            }

            // Получаем инициалы учителя из полного имени
            String teacherInitials = getTeacherInitials(teacher);

            log.debug("Сравнение инициалов: '{}' (из отчета) vs '{}' (из БД)",
                    initials, teacherInitials);

            return initials.equalsIgnoreCase(teacherInitials);

        } catch (Exception e) {
            log.warn("Ошибка при сопоставлении короткого имени '{}': {}", shortName, e.getMessage());
            return false;
        }
    }

    /**
     * Извлекает инициалы из короткого имени "Фамилия И.О."
     */
    private String extractInitialsFromShortName(String shortName) {
        StringBuilder initials = new StringBuilder();

        // Убираем фамилию (первое слово)
        String[] parts = shortName.split("\\s+");
        for (int i = 1; i < parts.length; i++) {
            String part = parts[i];
            // Берем первую букву (убираем точку если есть)
            if (!part.isEmpty()) {
                char firstChar = part.charAt(0);
                if (Character.isLetter(firstChar)) {
                    initials.append(Character.toLowerCase(firstChar));
                }
            }
        }

        return initials.toString();
    }

    /**
     * Получает инициалы учителя из полного имени
     */
    private String getTeacherInitials(Teacher teacher) {
        StringBuilder initials = new StringBuilder();

        if (teacher.getFirstName() != null && !teacher.getFirstName().isEmpty()) {
            initials.append(Character.toLowerCase(teacher.getFirstName().charAt(0)));
        }

        if (teacher.getMiddleName() != null && !teacher.getMiddleName().isEmpty()) {
            initials.append(Character.toLowerCase(teacher.getMiddleName().charAt(0)));
        }

        return initials.toString();
    }

    /**
     * Сопоставление при разном порядке слов
     */
    private boolean matchDifferentOrder(String reportNormalized, String dbNormalized,
                                        String[] reportParts, String[] dbParts) {
        // Создаем наборы слов (без фамилии)
        Set<String> reportWords = new HashSet<>();
        Set<String> dbWords = new HashSet<>();

        for (int i = 1; i < reportParts.length; i++) {
            if (reportParts[i].length() == 1) {
                // Это инициал
                reportWords.add(reportParts[i]);
            } else if (reportParts[i].length() > 1) {
                // Это имя/отчество полностью
                reportWords.add(reportParts[i]);
            }
        }

        for (int i = 1; i < dbParts.length; i++) {
            if (dbParts[i].length() == 1) {
                // Это инициал
                dbWords.add(dbParts[i]);
            } else if (dbParts[i].length() > 1) {
                // Это имя/отчество полностью
                dbWords.add(dbParts[i]);
            }
        }

        // Проверяем пересечение
        reportWords.retainAll(dbWords);

        if (!reportWords.isEmpty()) {
            log.debug("✓ Совпадение по словам (разный порядок): {}", reportWords);
            return true;
        }

        return false;
    }

    /**
     * Поиск учителя по всем возможным вариантам
     * @return Optional с найденным учителем
     */
    public Optional<Teacher> findMatchingTeacher(String teacherNameFromReport, List<Teacher> allTeachers) {
        if (teacherNameFromReport == null || teacherNameFromReport.trim().isEmpty() || allTeachers.isEmpty()) {
            return Optional.empty();
        }

        log.debug("Поиск учителя для: '{}' (всего учителей: {})",
                teacherNameFromReport, allTeachers.size());

        // Сначала пробуем точное совпадение
        for (Teacher teacher : allTeachers) {
            if (matchTeacher(teacherNameFromReport, teacher)) {
                log.debug("Найден учитель: '{}' для '{}'", teacher.getFullName(), teacherNameFromReport);
                return Optional.of(teacher);
            }
        }

        // Если точного нет, пробуем ослабленные варианты
        return findFuzzyMatch(teacherNameFromReport, allTeachers);
    }

    /**
     * Нечеткий поиск (ослабленные правила)
     */
    private Optional<Teacher> findFuzzyMatch(String teacherNameFromReport, List<Teacher> allTeachers) {
        String reportNormalized = normalizeForComparison(teacherNameFromReport);
        String[] reportParts = reportNormalized.split(" ");

        if (reportParts.length == 0) {
            return Optional.empty();
        }

        String reportSurname = reportParts[0];

        // Собираем кандидатов с такой же фамилией
        List<Teacher> candidates = new ArrayList<>();
        for (Teacher teacher : allTeachers) {
            String dbNormalized = normalizeForComparison(teacher.getFullName());
            String[] dbParts = dbNormalized.split(" ");

            if (dbParts.length > 0 && reportSurname.equals(dbParts[0])) {
                candidates.add(teacher);
            }
        }

        if (candidates.isEmpty()) {
            return Optional.empty();
        }

        // Если только один кандидат с такой фамилией - берем его
        if (candidates.size() == 1) {
            Teacher teacher = candidates.get(0);
            log.debug("Найден единственный кандидат с фамилией '{}': {}",
                    reportSurname, teacher.getFullName());
            return Optional.of(teacher);
        }

        // Если несколько кандидатов, пробуем уточнить по инициалам
        if (reportParts.length >= 2) {
            String reportInitials = extractInitialsFromParts(reportParts);

            for (Teacher teacher : candidates) {
                String teacherInitials = getTeacherInitials(teacher);
                if (reportInitials.equalsIgnoreCase(teacherInitials)) {
                    log.debug("Найден по инициалам: '{}' для '{}'",
                            teacher.getFullName(), teacherNameFromReport);
                    return Optional.of(teacher);
                }
            }
        }

        // Если ничего не нашли, возвращаем первого кандидата
        log.debug("Несколько кандидатов, берем первого: {}", candidates.get(0).getFullName());
        return Optional.of(candidates.get(0));
    }

    /**
     * Извлечение инициалов из частей имени
     */
    private String extractInitialsFromParts(String[] parts) {
        StringBuilder initials = new StringBuilder();

        for (int i = 1; i < parts.length && i < 3; i++) {
            if (!parts[i].isEmpty()) {
                initials.append(Character.toLowerCase(parts[i].charAt(0)));
            }
        }

        return initials.toString();
    }
}