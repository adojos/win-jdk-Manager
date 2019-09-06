
@echo off
echo Setting JAVA_HOME 9
setx JAVA_HOME C:\JDKs\jdk-9.0.4 /m
echo Setting JRE_HOME 9
setx JRE_HOME C:\JDKs\jre-9.0.4 /m
echo setting PATH For Java 9
setx PATH C:\JDKs\jdk-9.0.4\bin;%PATH% /m
echo Display java version
java -version
pause