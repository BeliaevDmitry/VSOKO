# Dockerfile
FROM openjdk:11-jdk-slim

# Устанавливаем необходимые утилиты
RUN apt-get update && apt-get install -y \
    maven \
    && rm -rf /var/lib/apt/lists/*

# Создаем рабочую директорию
WORKDIR /app

# Копируем исходный код
COPY pom.xml .
COPY src ./src

# Собираем проект
RUN mvn clean package -DskipTests

# Создаем папки для данных
RUN mkdir -p /data/input /data/output /data/processed /data/archive

# Копируем собранный JAR
RUN cp target/*.jar app.jar

# Указываем точку входа
ENTRYPOINT ["java", "-jar", "app.jar"]"]