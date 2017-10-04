var WshShell = WScript.CreateObject("WScript.Shell");
var WshFSO = WScript.CreateObject("Scripting.FileSystemObject");
CurDir = WshFSO.GetAbsolutePathName(WshShell.CurrentDirectory);

dbdirud = "D:\\1c_archiv\\1cpromud\\dumpbase";
cfdirud = "D:\\1c_archiv\\1cpromud\\dumpdbcfg";
dbdirtoud = "\\\\sirius\\backup$\\1c_archiv\\1cprom\\dumpbase";
cfdirtoud = "\\\\sirius\\backup$\\1c_archiv\\1cprom\\dumpdbcfg";

dbdirsi = "D:\\1c_archiv\\1cpromsi\\dumpbase";
cfdirsi = "D:\\1c_archiv\\1cpromsi\\dumpdbcfg";
dbdirtosi = "\\\\sirius\\backup$\\1c_archiv\\sovintel\\dumpbase";
cfdirtosi = "\\\\sirius\\backup$\\1c_archiv\\sovintel\\dumpdbcfg";

var LogFile;

var startDate = new Date();
startDate.setHours(startDate.getHours()+12);

var dt = new Date();
var MsgDate = String(dt.getYear()) + "-" +( (dt.getMonth() > 8) ? String(dt.getMonth()+1) : ("0" + String(dt.getMonth()+1)) ) + "-" + ((dt.getDate() > 9) ? String(dt.getDate()) : ("0"+String(dt.getDate())) );

var strLogFileName = CurDir + "\\LOGs\\Msg_" + MsgDate +".txt";	

// функци€ добавлени€ строки в лог
function LogWrite(aTextLine){
	var da = new Date();
	
	try{	
		if(aTextLine == "")
			LogFile.WriteBlankLines(1);
		else	
			LogFile.WriteLine(da.toLocaleString() + " : " + aTextLine);
	}catch(e)
	{
		// ошибка записи 
		return(0);
	}
	return(1);
} // end function LogWrite

function CloseLogFile(){
	try{ 
		LogFile.Close() 
	}catch(e){
	};	
} // end function CloseLogFile

function OpenLogFile(){
	WScript.Sleep(2000);
	try{
		LogFile = WshFSO.OpenTextFile(strLogFileName,8,1);
	}catch(e)
	{
		// ошибка создани€ текстового файла
	 }
} // end function OpenLogFile

srvnameUD = "epsilon\\wtc_work_81";
usernameUD = "јдминистратор";
userpassUD = "admin1csysx";
path_1cv8UD = "\"C:\\Program Files (x86)\\1cv81\\bin\\1cv8.exe\"";

srvnameSI = "epsilon\\Sovintel_81";
usernameSI = "јдминистратор";
userpassSI = "adminx";
path_1cv8SI = "\"C:\\Program Files (x86)\\1cv81\\bin\\1cv8.exe\"";

OpenLogFile();
LogWrite("");
LogWrite("***** START *****");
LogWrite("---------------------------------------------------------");


LogWrite("START Restarting 1Cv81 server...");

CloseLogFile();
WshShell.Run("cscript.exe " + CurDir + "\\RestartSrv_1Cv81.js",2,1);
OpenLogFile();

LogWrite("END Restarting 1C server");

LogWrite("begin copy wtc_work_81");
LogWrite("Dumping UD IB into " + dbdirud);
CloseLogFile();
var d = new Date();
var FDate = String(d.getYear()) + "-" + String(d.getMonth()+1) + "-" + ((d.getDate() > 9) ? String(d.getDate()) : ("0"+String(d.getDate())) );
WshShell.Run(path_1cv8UD + " CONFIG /S" + srvnameUD + " /N" + usernameUD + " /P" + userpassUD + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /DumpIB" + dbdirud + "\\IB" + FDate + ".dt",2,1);
OpenLogFile();
LogWrite("Dumping configuration into " + cfdirud);
CloseLogFile();
WshShell.Run(path_1cv8UD + " CONFIG /S" + srvnameUD + " /N" + usernameUD + " /P" + userpassUD + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /DumpCFG" + cfdirud + "\\CF" + FDate + ".cf",2,1);
OpenLogFile();
LogWrite("end copy wtc_work_81");


LogWrite("begin copy Sovintel_81");
LogWrite("Dumping SI IB into " + dbdirsi);
CloseLogFile();
var d = new Date();
var FDate = String(d.getYear()) + "-" + String(d.getMonth()+1) + "-" + ((d.getDate() > 9) ? String(d.getDate()) : ("0"+String(d.getDate())) );
WshShell.Run(path_1cv8SI + " CONFIG /S" + srvnameSI + " /N" + usernameSI + " /P" + userpassSI + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /DumpIB" + dbdirsi + "\\IB" + FDate + ".dt",2,1);
OpenLogFile();
LogWrite("Dumping configuration into " + cfdirsi);
CloseLogFile();
WshShell.Run(path_1cv8SI + " CONFIG /S" + srvnameSI + " /N" + usernameSI + " /P" + userpassSI + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /DumpCFG" + cfdirsi + "\\CF" + FDate + ".cf",2,1);
OpenLogFile();
LogWrite("end copy Sovintel_81");


LogWrite("START Restarting 1C server...");
CloseLogFile();
WshShell.Run("cscript.exe " + CurDir + "\\RestartSrv_1Cv81.js",2,1);
OpenLogFile();
LogWrite("END Restarting 1C server");

if((startDate.getDay()==6)||(startDate.getDay()==0)){
	LogWrite("SKIPED Updating DB UD configuration");
}else{

	LogWrite("START Updating DB UD configuration");
	CloseLogFile();
	WshShell.Run(path_1cv8UD + " CONFIG /S" + srvnameUD + " /N" + usernameUD + " /P" + userpassUD + " /UpdateDBCfg /Out" + strLogFileName + " -NoTruncate",2,1);
	OpenLogFile();
	LogWrite("END Updating DB UD configuration");
}

LogWrite("REINDEX  WTC_WORK_81");
CloseLogFile();
WshShell.Run(path_1cv8UD + " CONFIG /S" + srvnameUD + " /N" + usernameUD + " /P" + userpassUD + " /IBCheckAndRepair -Reindex /Out" + strLogFileName + " -NoTruncate",2,1);
OpenLogFile();
LogWrite("END REINDEX WTC_WORK_81");

if((startDate.getDay()==6)||(startDate.getDay()==0)){
	LogWrite("SKIPED Updating DB SI configuration");
}else{
	LogWrite("START Updating DB SI configuration");
	CloseLogFile();
	WshShell.Run(path_1cv8SI + " CONFIG /S" + srvnameSI + " /N" + usernameSI + " /P" + userpassSI + " /UpdateDBCfg /Out" + strLogFileName + " -NoTruncate",2,1);
	OpenLogFile();
	LogWrite("END Updating DB SI configuration");
}

LogWrite("REINDEX  Sovintel_81");
CloseLogFile();
WshShell.Run(path_1cv8SI + " CONFIG /S" + srvnameSI + " /N" + usernameSI + " /P" + userpassSI + " /IBCheckAndRepair -Reindex /Out" + strLogFileName + " -NoTruncate",2,1);
OpenLogFile();
LogWrite("END REINDEX SOVINTEL_81");


LogWrite("begin copy wtc_work_81, unload 1C base on disk Q");
try {
	WshFSO.CopyFile(dbdirud + "\\*.*", dbdirtoud, 1); // затереть старый, если есть
	WshFSO.DeleteFile(dbdirud + "\\*.*", 1);
}catch(e){
	// error
	LogWrite("error copying db files from " + dbdirud + " to " + dbdirtoud);
	
}

LogWrite("begin copy wtc_work_81 CFG, unload 1C conf on disk Q");
try {
	WshFSO.CopyFile(cfdirud + "\\*.*", cfdirtoud, 1); // затереть старый, если есть
	WshFSO.DeleteFile(cfdirud + "\\*.*", 1); 
}catch(e){
	// error
	LogWrite("error copying cf files from " + cfdirud + " to " + cfdirtoud);
}

LogWrite("end copy wtc_work_81");

LogWrite("---------------------------------------------------------");


LogWrite("begin copy Sovintel_81, unload 1C base on disk Q");
try {
	WshFSO.CopyFile(dbdirsi + "\\*.*", dbdirtosi, 1); // затереть старый, если есть
	WshFSO.DeleteFile(dbdirsi + "\\*.*", 1);
}catch(e){
	// error
	LogWrite("error copying db files from " + dbdirsi + " to " + dbdirtosi);
	
}

LogWrite("begin copy Sovintel_81 CFG, unload 1C conf on disk Q");
try {
	WshFSO.CopyFile(cfdirsi + "\\*.*", cfdirtosi, 1); // затереть старый, если есть
	WshFSO.DeleteFile(cfdirsi + "\\*.*", 1); 
}catch(e){
	// error
	LogWrite("error copying cf files from " + cfdirsi + " to " + cfdirtosi);
}

LogWrite("end copy Sovintel_81");

LogWrite("---------------------------------------------------------");


//LogWrite("START Unloading data from UD to FUS");
//CloseLogFile();
//WshShell.Run(path_1cv8UD + " ENTERPRISE /S" + srvnameUD + " /N" + usernameUD + " /P" + userpassUD + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /C\"/EP" + CurDir + "\\¬ыгрузка онтрагентовƒокументов¬‘”—.epf-ExitOnError\"",2,1);
//OpenLogFile();
//LogWrite("END   Unloading data from UD to FUS");

LogWrite("START Unloading data from UD to SI");
CloseLogFile();
WshShell.Run(path_1cv8UD + " ENTERPRISE /S" + srvnameUD + " /N" + usernameUD + " /P" + userpassUD + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /C\"/EP" + CurDir + "\\ќбмен”ƒ—овинтел.epf-ExitOnError\"",2,1);
OpenLogFile();
LogWrite("END   Unloading data from UD to SI");

LogWrite("START Unloading data from SI to UD");
CloseLogFile();
WshShell.Run(path_1cv8SI + " ENTERPRISE /S" + srvnameSI + " /N" + usernameSI + " /P" + userpassSI + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /C\"/EP" + CurDir + "\\ќбмен—овинтел”ƒ.epf-ExitOnError\"",2,1);
OpenLogFile();
LogWrite("END   Unloading data from SI to UD");

LogWrite("---------------------------------------------------------");

LogWrite("START Recalculate totals period in UD");
CloseLogFile();
WshShell.Run(path_1cv8UD + " ENTERPRISE /S" + srvnameUD + " /N" + usernameUD + " /P" + userpassUD + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /C\"/EP" + CurDir + "\\SetTotalsPeriod.epf-ExitOnError\"",2,1);
OpenLogFile();
LogWrite("END   Recalculate totals period in UD");

// Ѕƒѕ
LogWrite("---------------------------------------------------------");

LogWrite("START AutoUpdate dogovora from SI 6Q# to UD");
CloseLogFile();
WshShell.Run(path_1cv8UD + " ENTERPRISE /S" + srvnameUD + " /N" + usernameUD + " /P" + userpassUD + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /C\"/EP" + CurDir + "\\AutoUpdate_6Q_SI_UD.epf-ExitOnError\"",2,1);
OpenLogFile();
LogWrite("END AutoUpdate dogovora from SI 6Q# to UD");
// --- end LogWrite() 


LogWrite("START Recalculate totals period in SI");
CloseLogFile();
WshShell.Run(path_1cv8SI + " ENTERPRISE /S" + srvnameSI + " /N" + usernameSI + " /P" + userpassSI + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /C\"/EP" + CurDir + "\\SetTotalsPeriod.epf-ExitOnError\"",2,1);
OpenLogFile();
LogWrite("END   Recalculate totals period in SI");

//	LogWrite("START Changing cash desk numbers in UD");
//	CloseLogFile();
//	WshShell.Run(path_1cv8UD + " ENTERPRISE /S" + srvnameUD + " /N" + usernameUD + " /P" + userpassUD + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /C\"/EP" + CurDir + "\\SetKassa.epf-ExitOnError\"",2,1);
//	OpenLogFile();
//	LogWrite("END   Changing cash desk numbers in UD");

try {
	if(WshFSO.FolderExists("\\\\sirius\\backup$\\1c_archiv\\1cprom\\LOGs"))
		WshFSO.CopyFile(strLogFileName, "\\\\sirius\\backup$\\1c_archiv\\1cprom\\LOGs\\", 1); // затереть старый, если есть
	else{
		OpenLogFile();
		LogWrite("error: Folder \"\\\\sirius\\backup$\\1c_archiv\\1cprom\\LOGs\" not found!");
		CloseLogFile();
	};
}catch(e){
	// error
	OpenLogFile();
	LogWrite("error copying LOG");
	CloseLogFile();
}

OpenLogFile();

LogWrite("***** The END *****");
CloseLogFile();

WshShell.Run(CurDir + "\\DelFiles.exe \"\\\\sirius\\backup$\\1c_archiv\\1CPROM\\DUMPBASE\" 7 dt",2,1);
WshShell.Run(CurDir + "\\DelFiles.exe \"\\\\sirius\\backup$\\1c_archiv\\1CPROM\\DUMPDBCFG\" 7 cf",2,1);

WshShell.Run(CurDir + "\\DelFiles.exe \"\\\\sirius\\backup$\\1c_archiv\\Sovintel\\DUMPBASE\" 7 dt",2,1);
WshShell.Run(CurDir + "\\DelFiles.exe \"\\\\sirius\\backup$\\1c_archiv\\Sovintel\\DUMPDBCFG\" 7 cf",2,1);
