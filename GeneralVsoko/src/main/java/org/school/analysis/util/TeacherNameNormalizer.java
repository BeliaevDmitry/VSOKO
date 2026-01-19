package org.school.analysis.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.text.similarity.LevenshteinDistance;

import java.text.Normalizer;
import java.util.*;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

@Slf4j
public class TeacherNameNormalizer {

    private static final Pattern DIACRITICS = Pattern.compile("\\p{InCombiningDiacriticalMarks}+");
    private static final Pattern MULTIPLE_SPACES = Pattern.compile("\\s+");
    private static final Pattern PUNCTUATION = Pattern.compile("[.,!?;:()\\[\\]{}<>\\/\\\\\"'`~@#$%^&*_+=|]");
    private static final LevenshteinDistance LEVENSHTEIN = new LevenshteinDistance();

    // Типичные сокращения
    private static final Map<String, String> COMMON_SHORTENINGS = Stream.of(
            new String[][] {
                    {"алл.", "алла"},
                    {"андр.", "андрей"},
                    {"анн.", "анна"},
                    {"влад.", "владимир"},
                    {"дмитр.", "дмитрий"},
                    {"евг.", "евгений"},
                    {"макс.", "максим"},
                    {"мих.", "михаил"},
                    {"ник.", "николай"},
                    {"олег", "олег"},
                    {"петр", "петр"},
                    {"серг.", "сергей"},
                    {"тат.", "татьяна"},
                    {"юл.", "юлия"},
                    {"юр.", "юрий"},
                    {"и.", "иван"},
                    {"в.", "владимир"},
                    {"а.", "александр"},
                    {"с.", "сергей"}
            }).collect(Collectors.toMap(data -> data[0], data -> data[1]));

    /**
     * Нормализация имени для сравнения
     */
    public static String normalize(String name) {
        if (name == null || name.trim().isEmpty()) {
            return "";
        }

        String result = name.toLowerCase();

        // 1. Заменяем сокращения на полные формы
        result = expandShortenings(result);

        // 2. Удаляем пунктуацию
        result = PUNCTUATION.matcher(result).replaceAll("");

        // 3. Заменяем все последовательности пробелов одним пробелом
        result = MULTIPLE_SPACES.matcher(result.trim()).replaceAll(" ");

        // 4. Нормализуем русские символы (латинские -> русские)
        result = normalizeRussian(result);

        // 5. Обработка инициалов
        result = normalizeInitials(result);

        // 6. Убираем диакритические знаки (если есть)
        result = DIACRITICS.matcher(Normalizer.normalize(result, Normalizer.Form.NFD))
                .replaceAll("");

        // 7. Сортируем части ФИО (фамилия всегда первая)
        result = sortNameParts(result);

        return result.trim();
    }

    /**
     * Нормализация русских символов (замена латинских на русские)
     */
    private static String normalizeRussian(String text) {
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
     * Нормализация инициалов
     * Примеры:
     * - "и.и." → "и и"
     * - "и . и ." → "и и"
     * - "и  и  " → "и и"
     */
    private static String normalizeInitials(String text) {
        // Удаляем точки возле букв
        String result = text.replaceAll("([а-я])\\.", "$1");

        // Заменяем инициалы с точками
        String[] parts = result.split("\\s+");
        List<String> normalizedParts = new ArrayList<>();

        for (String part : parts) {
            if (part.length() == 1 && Character.isLetter(part.charAt(0))) {
                // Одиночная буква - это инициал
                normalizedParts.add(part);
            } else if (part.length() > 1) {
                // Слово целиком
                normalizedParts.add(part);
            }
        }

        return String.join(" ", normalizedParts);
    }

    /**
     * Расширение сокращений
     */
    private static String expandShortenings(String text) {
        String result = text;
        for (Map.Entry<String, String> entry : COMMON_SHORTENINGS.entrySet()) {
            // Используем границы слов для точного совпадения
            String pattern = "\\b" + Pattern.quote(entry.getKey()) + "\\b";
            result = result.replaceAll(pattern, entry.getValue());
        }
        return result;
    }

    /**
     * Сортировка частей имени для унификации
     * Фамилия всегда идет первой, затем имя, затем отчество
     */
    private static String sortNameParts(String name) {
        String[] parts = name.split("\\s+");
        if (parts.length <= 1) {
            return name;
        }

        // Определяем фамилию (самая длинная часть или первая)
        String surname = parts[0];
        if (parts.length > 1) {
            for (int i = 1; i < parts.length; i++) {
                if (parts[i].length() > surname.length() && parts[i].length() > 2) {
                    surname = parts[i];
                }
            }
        }

        List<String> otherParts = new ArrayList<>();
        for (String part : parts) {
            if (!part.equals(surname)) {
                otherParts.add(part);
            }
        }

        // Сортируем остальные части по длине (имя обычно короче отчества)
        otherParts.sort(Comparator.comparingInt(String::length));

        // Собираем: фамилия + остальные части
        List<String> result = new ArrayList<>();
        result.add(surname);
        result.addAll(otherParts);

        return String.join(" ", result);
    }

    /**
     * Получение короткого имени (Фамилия И.О.)
     */
    public static String getShortName(String fullName) {
        String normalized = normalize(fullName);
        String[] parts = normalized.split("\\s+");

        if (parts.length < 2) {
            return capitalizeFirst(fullName.trim());
        }

        StringBuilder shortName = new StringBuilder(capitalizeFirst(parts[0])); // Фамилия

        for (int i = 1; i < Math.min(parts.length, 3); i++) {
            if (!parts[i].isEmpty()) {
                shortName.append(" ")
                        .append(Character.toUpperCase(parts[i].charAt(0)))
                        .append(".");
            }
        }

        return shortName.toString();
    }

    /**
     * Проверка схожести имен
     */
    public static boolean isSimilar(String name1, String name2) {
        if (name1 == null || name2 == null) {
            return false;
        }

        String norm1 = normalize(name1);
        String norm2 = normalize(name2);

        // Быстрое сравнение
        if (norm1.equals(norm2)) {
            return true;
        }

        // Проверка по Левенштейну
        Integer distance = LEVENSHTEIN.apply(norm1, norm2);
        int maxLength = Math.max(norm1.length(), norm2.length());

        if (maxLength == 0) {
            return false;
        }

        double similarity = 1.0 - (double) distance / maxLength;

        // Для коротких имен (с инициалами) требуем более высокую схожесть
        if (norm1.length() < 10 || norm2.length() < 10) {
            return similarity >= 0.8; // 80% для коротких имен
        }

        return similarity >= 0.7; // 70% для полных имен
    }

    /**
     * Проверка совпадения инициалов
     */
    public static boolean checkInitialsMatch(String name1, String name2) {
        String norm1 = normalize(name1);
        String norm2 = normalize(name2);

        String[] parts1 = norm1.split("\\s+");
        String[] parts2 = norm2.split("\\s+");

        if (parts1.length <= 1 || parts2.length <= 1) {
            return false; // Нет инициалов для сравнения
        }

        // Проверяем совпадение фамилии
        if (!parts1[0].equals(parts2[0])) {
            return false;
        }

        // Проверяем совпадение инициалов
        for (int i = 1; i < Math.min(parts1.length, parts2.length); i++) {
            if (parts1[i].length() > 0 && parts2[i].length() > 0) {
                if (parts1[i].charAt(0) != parts2[i].charAt(0)) {
                    return false;
                }
            }
        }

        return true;
    }

    /**
     * Извлечение фамилии из строки
     */
    public static String extractLastName(String name) {
        String normalized = normalize(name);
        String[] parts = normalized.split("\\s+");

        if (parts.length == 0) {
            return "";
        }

        return capitalizeFirst(parts[0]);
    }

    /**
     * Извлечение инициалов из строки
     */
    public static String extractInitials(String name) {
        String normalized = normalize(name);
        String[] parts = normalized.split("\\s+");

        StringBuilder initials = new StringBuilder();
        for (int i = 1; i < parts.length && i < 3; i++) {
            if (parts[i].length() > 0) {
                initials.append(Character.toUpperCase(parts[i].charAt(0)))
                        .append(".");
            }
        }

        return initials.toString();
    }

    /**
     * Приведение первой буквы к заглавной
     */
    public static String capitalizeFirst(String word) {
        if (word == null || word.isEmpty()) {
            return word;
        }
        return Character.toUpperCase(word.charAt(0)) + word.substring(1);
    }

    /**
     * Приведение имени к нормальному виду (с заглавными буквами)
     */
    public static String capitalizeName(String name) {
        if (name == null || name.isEmpty()) {
            return name;
        }

        String normalized = normalize(name);
        String[] parts = normalized.split("\\s+");
        StringBuilder result = new StringBuilder();

        for (String part : parts) {
            if (!part.isEmpty()) {
                if (result.length() > 0) {
                    result.append(" ");
                }
                result.append(capitalizeFirst(part));
            }
        }

        return result.toString();
    }

    /**
     * Дополнительный метод для строгой проверки формата "Фамилия И.О."
     */
    public static boolean isValidShortNameFormat(String name) {
        if (name == null) {
            return false;
        }

        String normalized = normalize(name);
        String[] parts = normalized.split("\\s+");

        // Должно быть как минимум 2 части (фамилия и один инициал)
        if (parts.length < 2) {
            return false;
        }

        // Первая часть (фамилия) должна быть длиннее 1 символа
        if (parts[0].length() <= 1) {
            return false;
        }

        // Остальные части должны быть одиночными буквами (инициалы)
        for (int i = 1; i < parts.length; i++) {
            if (parts[i].length() != 1 || !Character.isLetter(parts[i].charAt(0))) {
                return false;
            }
        }

        return true;
    }
}