FROM openjdk:8-jre-alpine
WORKDIR /root/app
RUN apk add openjdk8
COPY . /root/app
ENTRYPOINT ["java", "-jar", "/root/app/target/DiffExcel-1.jar"]
CMD ["/root/app/db_2020.xlsx", "/root/app/db_2019.xlsx"]
