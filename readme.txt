src/main/java/org/school/analysis/
├── 📂 model/                          # Модели данных (DTO/Entity)
│   ├── StudentResult.java            # Результат ученика
│   ├── ReportFile.java               # Файл отчета + метаданные
│   ├── ParseResult.java              # Результат парсинга
│   ├── ProcessingStatus.java         # Enum статусов обработки
│   └── TestMetadata.java             # Метаданные теста
│
├── 📂 service/                        # Сервисный слой (бизнес-логика)
│   ├── 📂 impl/                       # Реализации сервисов
│   │   ├── ReportProcessorServiceImpl.java     # Главный координатор
│   │   ├── ReportParserServiceImpl.java        # Парсинг Excel (с заменой ExcelReportParser)
│   │   └── FileOrganizerServiceImpl.java       # Организация файлов
│   │
│   ├── ReportProcessorService.java             # Интерфейс главного сервиса
│   ├── ReportParserService.java                # Интерфейс парсинга
│   └── FileOrganizerService.java               # Интерфейс организации файлов
│
├── 📂 parser/                         # Парсеры разных форматов
│   └── 📂 strategy/                   # Стратегии парсинга
│       ├── StudentDataParser.java     # Парсинг данных учеников
│       └── MetadataParser.java        # Парсинг метаданных
│
├── 📂 repository/                        # репозиторий
│   ├── 📂 impl/                       # Реализация репозитория
│   │   ├── StudentResultRepositoryImpl.java     # Главный репозиторий
│   │
│   ├── ReportFileRepository.java             # репозиторий JPA
│   ├── StudentResultRepository.java          # репозиторий JPA
│
├── 📂 util/                           # Утилиты и хелперы
│   ├── ExcelParser.java               # Низкоуровневый парсинг Excel
│   └── ValidationHelper.java          # Валидация данных
│
├── 📂 config/                         # Конфигурация
│   └── AppConfig.java                 # Настройки путей и параметров
│
├── 📂 exception/                      # Кастомные исключения (НОВОЕ)
│   └── ValidationException.java       # Исключение для ошибок валидации
│
└── Main.java                          # Точка входа