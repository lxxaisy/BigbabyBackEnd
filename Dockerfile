# 该镜像需要依赖的基础镜像
FROM openjdk:11-jdk-slim
# 声明服务运行在8080端口
EXPOSE 8080
VOLUME /tmp
# 将jar包添加到容器中并更名为app.jar
COPY target/*.jar app.jar
# 执行命令
ENTRYPOINT ["java","-Djava.security.egd=file:/dev/./urandom","-jar","/app.jar"]