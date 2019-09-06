
@echo off
echo Setting JAVA_HOME 8
setx JAVA_HOME C:\JDKs\jdk1.8.0_212 /m
echo Setting JRE_HOME 8
setx JRE_HOME C:\JDKs\jre1.8.0_212 /m
echo setting PATH For Java 8
setx PATH C:\JDKs\jdk1.8.0_212\bin;%PATH% /m
echo Display java version
java -version
pause
