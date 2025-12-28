package org.school.analysis.repository;

import org.school.analysis.model.StudentResult;

import java.util.List;

/**
 * Репозиторий для сохранения результатов в БД
 */
public interface StudentResultRepository {

    /**
     * Сохранить результаты одного ученика
     */
    boolean save(StudentResult studentResult);

    /**
     * Сохранить список результатов
     */
    int saveAll(List<StudentResult> studentResults);

    /**
     * Проверить существование ученика в БД
     */
    boolean exists(StudentResult studentResult);
}