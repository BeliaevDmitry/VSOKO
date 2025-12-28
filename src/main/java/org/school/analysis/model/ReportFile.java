package org.school.analysis.model;

import lombok.Data;
import java.io.File;
import java.time.LocalDateTime;

@Data
public class ReportFile {
    private File file;                    // Файл
    private String subject;               // Предмет (из файла или имени)
    private String className;             // Класс (из файла или имени)
    private ProcessingStatus status;      // Статус обработки
    private LocalDateTime processedAt;    // Когда обработан
    private String errorMessage;          // Сообщение об ошибке (если есть)
    private int studentCount;             // Количество учеников в файле
}