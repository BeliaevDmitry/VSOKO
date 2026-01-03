-- 1. Создайте таблицу report_files
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
    max_total_score INTEGER,
    test_type VARCHAR(50),
    comment TEXT,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
);

-- 2. Создайте таблицу report_file_max_scores
CREATE TABLE IF NOT EXISTS report_file_max_scores (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    report_file_id UUID NOT NULL,
    task_number INTEGER NOT NULL,
    max_score INTEGER NOT NULL,
    CONSTRAINT fk_report_file_max_scores_report_file
        FOREIGN KEY (report_file_id)
        REFERENCES report_files(id)
        ON DELETE CASCADE
);

-- 3. Создайте таблицу student_results
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
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT fk_student_results_report_file
        FOREIGN KEY (report_file_id)
        REFERENCES report_files(id)
        ON DELETE CASCADE
);

-- 4. Создайте таблицу student_task_scores
CREATE TABLE IF NOT EXISTS student_task_scores (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    student_result_id UUID NOT NULL,
    task_number INTEGER NOT NULL,
    score INTEGER NOT NULL,
    max_score INTEGER NOT NULL,
    CONSTRAINT fk_student_task_scores_student_result
        FOREIGN KEY (student_result_id)
        REFERENCES student_results(id)
        ON DELETE CASCADE
);

-- 5. Создайте индексы для производительности
CREATE INDEX IF NOT EXISTS idx_report_files_file_hash ON report_files(file_hash);
CREATE INDEX IF NOT EXISTS idx_report_files_status ON report_files(status);
CREATE INDEX IF NOT EXISTS idx_report_files_subject_class ON report_files(subject, class_name);

CREATE INDEX IF NOT EXISTS idx_report_file_max_scores_report_file_id ON report_file_max_scores(report_file_id);
CREATE INDEX IF NOT EXISTS idx_report_file_max_scores_task_number ON report_file_max_scores(task_number);

CREATE INDEX IF NOT EXISTS idx_student_results_report_file_id ON student_results(report_file_id);
CREATE INDEX IF NOT EXISTS idx_student_results_fio ON student_results(fio);
CREATE INDEX IF NOT EXISTS idx_student_results_subject_class ON student_results(subject, class_name);

CREATE INDEX IF NOT EXISTS idx_student_task_scores_student_result_id ON student_task_scores(student_result_id);
CREATE INDEX IF NOT EXISTS idx_student_task_scores_task_number ON student_task_scores(task_number);