-- 1. Таблица report_files - ДОБАВЛЯЕМ колонку для JSON, УДАЛЯЕМ max_total_score
CREATE TABLE IF NOT EXISTS report_files (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    file_path VARCHAR(500) NOT NULL,
    file_name VARCHAR(255) NOT NULL,
    file_hash VARCHAR(64) NOT NULL UNIQUE,
    subject VARCHAR(100) NOT NULL,
    class_name VARCHAR(50) NOT NULL,
    status VARCHAR(50) NOT NULL,
    processed_at TIMESTAMP,
    error_message TEXT,
    student_count INTEGER DEFAULT 0,
    test_date DATE,
    teacher VARCHAR(200),
    school VARCHAR(200) DEFAULT 'ГБОУ №7',
    task_count INTEGER,
    -- УДАЛЕНО: max_total_score INTEGER, (будем вычислять из JSON)
    test_type VARCHAR(50),
    comment TEXT,
    -- НОВАЯ колонка для хранения максимальных баллов как JSON
    max_scores_json TEXT,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
);

-- 2. Таблицу report_file_max_scores УДАЛЯЕМ (больше не нужна)
-- DROP TABLE IF EXISTS report_file_max_scores;

-- 3. Таблица student_results
CREATE TABLE IF NOT EXISTS student_results (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    report_file_id UUID NOT NULL,
    subject VARCHAR(100) NOT NULL,
    class_name VARCHAR(50) NOT NULL,
    fio VARCHAR(200) NOT NULL,
    presence VARCHAR(50) NOT NULL,
    variant VARCHAR(100),
    test_type VARCHAR(50),
    test_date DATE NOT NULL,
    total_score INTEGER,
    percentage_score DOUBLE PRECISION,
    -- Колонка для хранения баллов учеников как JSON
    task_scores_json TEXT,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT fk_student_results_report_file
        FOREIGN KEY (report_file_id)
        REFERENCES report_files(id)
        ON DELETE CASCADE
);

-- 4. Создайте индексы для производительности
CREATE INDEX IF NOT EXISTS idx_report_files_file_hash ON report_files(file_hash);
CREATE INDEX IF NOT EXISTS idx_report_files_status ON report_files(status);
CREATE INDEX IF NOT EXISTS idx_report_files_subject_class ON report_files(subject, class_name);
CREATE INDEX IF NOT EXISTS idx_report_files_test_date ON report_files(test_date);

-- 5. Индексы для student_results
CREATE INDEX IF NOT EXISTS idx_student_results_report_file_id ON student_results(report_file_id);
CREATE INDEX IF NOT EXISTS idx_student_results_fio ON student_results(fio);
CREATE INDEX IF NOT EXISTS idx_student_results_subject_class ON student_results(subject, class_name);
CREATE INDEX IF NOT EXISTS idx_student_results_test_date ON student_results(test_date);

-- 6. Индексы для поиска по JSON (если используете PostgreSQL)
-- Для max_scores_json (опционально)
CREATE INDEX IF NOT EXISTS idx_report_files_max_scores_json
ON report_files USING gin (max_scores_json jsonb_path_ops);

-- Для task_scores_json
CREATE INDEX IF NOT EXISTS idx_student_results_task_scores_json
ON student_results USING gin (task_scores_json jsonb_path_ops);

-- 7. Дополнительные индексы для аналитики
CREATE INDEX IF NOT EXISTS idx_student_results_presence ON student_results(presence);
CREATE INDEX IF NOT EXISTS idx_student_results_total_score ON student_results(total_score);