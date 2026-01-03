# Dockerfile для вашего Java приложения
FROM eclipse-temurin:17-jdk-alpine

# 1. Устанавливаем рабочую директорию
WORKDIR /app

# 2. Копируем собранный JAR
COPY target/school-analysis-1.0.jar app.jar

# 3. Настраиваем запуск
ENTRYPOINT ["java", "-jar", "app.jar"]