📁 org.school.analysis/
├── 📂 model/                          # Модели данных
│   ├── 📂 dto/                        # DTO для передачи данных между слоями
│   │   ├── TestSummaryDto.java        # Сводные данные по тесту ✓
│   │   ├── StudentDetailedResultDto.java  # Детальные результаты студента ✓
│   │   ├── TaskStatisticsDto.java     # Статистика по заданию ✓
│   │   └── TeacherTestDetailDto.java  # Детальные данные теста для учителя ✓
│   │
│   ├── 📂 entity/                     # JPA сущности
│   │   ├── ReportFileEntity.java      # Сущность файла отчета
│   │   └── StudentResultEntity.java   # Сущность результата студента
│   │
│   ├── StudentResult.java             # Результат ученика ✓
│   ├── ReportFile.java               # Файл отчета + метаданные ✓
│   ├── ParseResult.java              # Результат парсинга ✓
│   ├── ProcessingStatus.java         # Enum статусов обработки ✓
│   └── TestMetadata.java             # Метаданные теста ✓
│
├── 📂 service/                        # Сервисный слой
│   ├── 📂 impl/                       # Реализации сервисов
│   │   ├── 📂 report/                 # Генерация отчетов
│   │   │   ├── ExcelReportServiceImpl.java      # Основной координатор
│   │   │   ├── ExcelReportBase.java             # Базовые методы и стили
│   │   │   ├── SummaryReportGenerator.java      # Сводные отчеты
│   │   │   ├── DetailReportGenerator.java       # Детальные отчеты
│   │   │   ├── TeacherReportGenerator.java      # Отчеты учителей
│   │   │   └── 📂 charts/            # Графики
│   │   │       ├── ExcelChartService.java       # Координатор графиков
│   │   │       ├── ExcelChartBase.java          # Базовый класс графиков
│   │   │       ├── ChartStyleConfig.java        # Конфигурация стилей
│   │   │       ├── StackedBarChartGenerator.java # Stacked диаграммы
│   │   │       ├── PercentageBarChartGenerator.java # Bar диаграммы
│   │   │       └── LineChartGenerator.java      # Line диаграммы
│   │   │
│   │   ├── GeneralServiceImpl.java           # Главный координатор ✓
│   │   ├── ParserServiceImpl.java           # Парсинг отчетов ✓
│   │   ├── FileOrganizerServiceImpl.java    # Организация файлов ✓
│   │   ├── SavedServiceImpl.java           # Сохранение в БД ✓
│   │   ├── AnalysisServiceImpl.java        # Анализ и статистика ✓
│   │
│   ├── GeneralService.java           # Интерфейс главного сервиса
│   ├── ParserService.java           # Интерфейс парсинга
│   ├── FileOrganizerService.java    # Интерфейс организации файлов
│   ├── SavedService.java           # Интерфейс сохранения в БД
│   ├── ExcelReportService.java     # Интерфейс генерации отчетов Excel
│   └── AnalysisService.java        # Интерфейс анализа данных
│
├── 📂 repository/                    # JPA репозитории
│   ├── ReportFileRepository.java    # Репозиторий для файлов отчетов
│   └── StudentResultRepository.java # Репозиторий для результатов студентов
│
├── 📂 parser/                        # Логика парсинга Excel
│   ├── 📂 strategy/                  # Стратегии парсинга
│   │   ├── MetadataParser.java      # Парсер метаданных ✓
│   │   └── StudentDataParser.java   # Парсер данных студентов ✓
│
├── 📂 util/                          # Утилиты и хелперы
│   ├── JsonScoreUtils.java          # Работа с JSON баллами ✓
│   ├── ValidationHelper.java        # Валидация данных
│   ├── DateTimeFormatters.java      # Форматтеры даты/времени ✓
│   └── ExcelUtils.java             # Общие утилиты для Excel
│
├── 📂 config/                        # Конфигурация
│   └── AppConfig.java               # Настройки путей и параметров ✓
│
├── 📂 exception/                     # Кастомные исключения
│   └── ValidationException.java     # Исключение для ошибок валидации ✓
│
├── 📂 mapper/                        # Мапперы для преобразования
│   └── ReportMapper.java            # Маппер между моделями и сущностями ✓
│
├── Main.java                 # Точка входа (Spring Boot)
