echo 正在设置临时环境变量
set JAVA_HOME=%~dp0source\jre7

start %JAVA_HOME%\bin\java -jar InsertPic2Doc.jar