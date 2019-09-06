
@echo off
echo Setting JAVA_HOME 10
setx JAVA_HOME C:\JDKs\jdk-10.0.2 /m
echo Setting JRE_HOME 10
setx JRE_HOME C:\JDKs\jre-10.0.2 /m
echo setting PATH For Java 8 10
setx PATH C:\JDKs\jdk-10.0.2\bin;%PATH% /m
echo Display java version
java -version
pause