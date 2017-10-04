// загрузка данных из OPERA

var WshShell = WScript.CreateObject("WScript.Shell");

//WScript.Interactive = 0;

path_1cv8KU = "\"C:\\Program Files (x86)\\1cv8\\8.3.8.1675\\bin\\1cv8.exe\"";
srvnameKU = "grand:3041\\wtc_work_ku";
usernameKU = "Рарус";
userpassKU = "rarususer";


strRun = path_1cv8KU + " ENTERPRISE /S" + srvnameKU + " /N" + usernameKU + " /P" + userpassKU + " /AU- /DisableStartupMessages /C\"2\"";

WshShell.Run(strRun,2,1);
