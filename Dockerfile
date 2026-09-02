# =========================
# BUILD
# =========================
FROM maven:3.9.11-eclipse-temurin-17 AS build

WORKDIR /app

COPY pom.xml .
RUN mvn dependency:go-offline -B

COPY src ./src

RUN mvn clean package -DskipTests


# =========================
# RUNTIME
# =========================
FROM eclipse-temurin:17-jre

WORKDIR /app

# JasperReports / PDF могут требовать системные шрифты
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
    fontconfig \
    fonts-dejavu-core \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

COPY --from=build /app/target/*.jar app.jar

EXPOSE 8080

ENV JAVA_OPTS="-Djava.awt.headless=true"

ENTRYPOINT ["sh", "-c", "java $JAVA_OPTS -jar app.jar"]