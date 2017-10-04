var WshShell = WScript.CreateObject("WScript.Shell");
var WshFSO = WScript.CreateObject("Scripting.FileSystemObject");
var WshNetwork = WScript.CreateObject("WScript.Network");

var colDrives = WshNetwork.EnumNetworkDrives(); 
if (colDrives.length != 0) { 
		
	for (i = 0; i < colDrives.length; i += 2) { 
		if(colDrives(i)=="T:")WshNetwork.RemoveNetworkDrive("T:",1,1);  
	} 	
}
WshNetwork.MapNetworkDrive("T:","\\\\fiona\\Departments\\ДЭиП\\Cognos\\Transfer",0);

CurDir = WshFSO.GetAbsolutePathName(WshShell.CurrentDirectory);

dbdir = "D:\\1c_archiv\\1cpromku\\dumpbase";
cfdir = "D:\\1c_archiv\\1cpromku\\dumpdbcfg";
dbdirto = "\\\\sirius\\backup$\\1c_archiv\\1cpromku\\dumpbase";
cfdirto = "\\\\sirius\\backup$\\1c_archiv\\1cpromku\\dumpdbcfg";

var LogFile;

var startDate = new Date();
startDate.setHours(startDate.getHours()+12);

var dt = new Date();
var MsgDate = String(dt.getYear()) + "-" +( (dt.getMonth() > 8) ? String(dt.getMonth()+1) : ("0" + String(dt.getMonth()+1)) ) + "-" + ((dt.getDate() > 9) ? String(dt.getDate()) : ("0"+String(dt.getDate())) );

var strLogFileName = CurDir + "\\LOGs\\Msg_" + MsgDate +".txt";	

// функция добавления строки в лог
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
		// ошибка создания текстового файла
	 }
} // end function OpenLogFile

srvnameKU = "kappa\\wtc_work_ku";
usernameKU = "OBMEN";
userpassKU = "Obmen2014";
usernameKUR = "Рарус";
userpassKUR = "rarususer";
path_1cv8KU = "\"C:\\Program Files (x86)\\1cv8\\8.3.9.2170\\bin\\1cv8.exe\"";

srvnameUD = "epsilon\\wtc_work_81";
usernameUD = "Администратор";
userpassUD = "admin1csysx";
path_1cv8UD = "\"C:\\Program Files (x86)\\1cv81\\bin\\1cv8.exe\"";

srvnameSI = "epsilon\\Sovintel_81";
usernameSI = "Администратор";
userpassSI = "adminx";
path_1cv8SI = "\"C:\\Program Files (x86)\\1cv81\\bin\\1cv8.exe\"";

OpenLogFile();
LogWrite("");
LogWrite("***** START *****");
LogWrite("---------------------------------------------------------");


//LogWrite("START Restarting 1Cv82 server...");
//
//CloseLogFile();
//WshShell.Run("cscript.exe " + CurDir + "\\RestartSrv_1Cv82.js",2,1);
//OpenLogFile();
//
//LogWrite("END Restarting 1C server");

//LogWrite("begin copy wtc_work_ku_82");
//LogWrite("SKIPED!!! NO Dumping KU IB into " + dbdir);
//CloseLogFile();
//var d = new Date();
//var FDate = String(d.getYear()) + "-" + String(d.getMonth()+1) + "-" + ((d.getDate() > 9) ? String(d.getDate()) : ("0"+String(d.getDate())) );

// WshShell.Run(path_1cv8KU + " CONFIG /S" + srvnameKU + " /N" + usernameKU + " /P" + userpassKU + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /DumpIB" + dbdir + "\\IB" + FDate + ".dt",2,1);

//OpenLogFile();
//LogWrite("Dumping configuration into " + cfdir);
//CloseLogFile();
//WshShell.Run(path_1cv8KU + " CONFIG /S" + srvnameKU + " /N" + usernameKU + " /P" + userpassKU + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /DumpCFG" + cfdir + "\\CF" + FDate + ".cf",2,1);
//OpenLogFile();
//LogWrite("end copy wtc_work_ku");


//LogWrite("START Restarting 1C 82 server...");
//CloseLogFile();
//WshShell.Run("cscript.exe " + CurDir + "\\RestartSrv_1Cv82.js",2,1);
//OpenLogFile();
//LogWrite("END Restarting 1C server");


LogWrite("START Terminating KU Sessions");
CloseLogFile();
WshShell.Run(path_1cv8KU + " ENTERPRISE /S" + srvnameKU + " /N" + usernameKUR + " /P" + userpassKUR + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /C\"6\"",2,1);
OpenLogFile();
LogWrite("END Terminating KU Sessions");


//if((startDate.getDay()==6)||(startDate.getDay()==0)){
//	LogWrite("-------SKIPED Updating DB KU configuration");
//}else{
	LogWrite("START Updating DB KU configuration");
	//LogWrite("-------SKIPPED!!!!! for the NY Holidays!!!!!");
	CloseLogFile();
	WshShell.Run(path_1cv8KU + " CONFIG /S" + srvnameKU + " /N" + usernameKU + " /P" + userpassKU + " /UpdateDBCfg /Out" + strLogFileName + " -NoTruncate",2,1);
	OpenLogFile();
	LogWrite("END Updating DB KU configuration");
//}

//LogWrite("START Restarting 1C server...");
//CloseLogFile();
//WshShell.Run("cscript.exe " + CurDir + "\\RestartSrv_1Cv82.js",2,1);
//OpenLogFile();
//LogWrite("END Restarting 1C server");

LogWrite("REINDEX  WTC_WORK_KU_83");
//LogWrite("-------SKIPPED!!!!! for the NY Holidays!!!!!");
CloseLogFile();
WshShell.Run(path_1cv8KU + " CONFIG /S" + srvnameKU + " /N" + usernameKU + " /P" + userpassKU + " /IBCheckAndRepair -Reindex /Out" + strLogFileName + " -NoTruncate",2,1);
OpenLogFile();
LogWrite("END REINDEX WTC_WORK_KU_83");


//LogWrite("begin copy wtc_work_ku_82, unload 1C base on disk Q");
//try {
//	WshFSO.CopyFile(dbdir + "\\*.*", dbdirto, 1); // затереть старый, если есть
//	WshFSO.DeleteFile(dbdir + "\\*.*", 1);
//}catch(e){
//	// error
//	LogWrite("error copying db files from " + dbdir + " to " + dbdirto);
//	
//}
//
//LogWrite("begin copy wtc_work_ku_82 CFG, unload 1C base on disk Q");
//try {
//	WshFSO.CopyFile(cfdir + "\\*.*", cfdirto, 1); // затереть старый, если есть
//	WshFSO.DeleteFile(cfdir + "\\*.*", 1); 
//}catch(e){
//	// error
//	LogWrite("error copying cf files from " + cfdir + " to " + cfdirto);
//}
//
//LogWrite("end copy wtc_work_ku_82");

LogWrite("---------------------------------------------------------");

//--------------------------------------------------------------------------------
// загрузка данных из БОсс - кадровик

LogWrite("START Загрузка из БОСС-Кадровик");
CloseLogFile();

WshShell.Run(path_1cv8KU + " ENTERPRISE /S" + srvnameKU + " /N" + usernameKUR + " /P" + userpassKUR + " /AU- /DisableStartupMessages /C\"5\"",2,1);
OpenLogFile();
LogWrite("END Загрузка из БОСС-Кадровик");
//--------------------------------------------------------------------------------

LogWrite("START Unloading data from UD");
CloseLogFile();
WshShell.Run(path_1cv8UD + " ENTERPRISE /S" + srvnameUD + " /N" + usernameUD + " /P" + userpassUD + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /C\"/EP" + CurDir + "\\Unload_UD_KU.epf-ExitOnError\"",2,1);
OpenLogFile();
LogWrite("END   Unloading data from UD");

LogWrite("START Unloading data from SI");
CloseLogFile();
WshShell.Run(path_1cv8SI + " ENTERPRISE /S" + srvnameSI + " /N" + usernameSI + " /P" + userpassSI + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /C\"/EP" + CurDir + "\\Unload_SI_KU.epf-ExitOnError\"",2,1);
OpenLogFile();
LogWrite("END   Unloading data from SI");

LogWrite("START Loading data from UD, SI unloading from KU");
CloseLogFile();
WshShell.Run(path_1cv8KU + " ENTERPRISE /S" + srvnameKU + " /N" + usernameKU + " /P" + userpassKU + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages",2,1);
OpenLogFile();
LogWrite("END   Loading data from UD, SI unloading from KU");

LogWrite("START Loading data from KU to UD");
CloseLogFile();
WshShell.Run(path_1cv8UD + " ENTERPRISE /S" + srvnameUD + " /N" + usernameUD + " /P" + userpassUD + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /C\"/EP" + CurDir + "\\Load_KU_UD.epf-ExitOnError\"",2,1);
OpenLogFile();
LogWrite("END Loading data from KU to UD");

LogWrite("START Loading data from KU to SI");
CloseLogFile();
WshShell.Run(path_1cv8SI + " ENTERPRISE /S" + srvnameSI + " /N" + usernameSI + " /P" + userpassSI + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /C\"/EP" + CurDir + "\\Load_KU_SI.epf-ExitOnError\"",2,1);
OpenLogFile();
LogWrite("END Loading data from KU to SI");

LogWrite("START Catalogue comparison in UD");
CloseLogFile();
WshShell.Run(path_1cv8UD + " ENTERPRISE /S" + srvnameUD + " /N" + usernameUD + " /P" + userpassUD + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /C\"/EP" + CurDir + "\\CatalogueAutoComparison.epf-ExitOnError\"",2,1);
OpenLogFile();
LogWrite("END Catalogue comparison in UD");

LogWrite("END OF Data exchanging");

LogWrite("!!! START OF EPF !!!");
usernameKU = "Рарус";
userpassKU = "rarususer";
LogWrite("---------------------------------------------------------");
CloseLogFile();
WshShell.Run(path_1cv8KU + " ENTERPRISE /S" + srvnameKU + " /N" + usernameKU + " /P" + userpassKU + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages",2,1);

OpenLogFile();
LogWrite("Start UD Cognos Processing");
CloseLogFile();
WshShell.Run(path_1cv8UD + " ENTERPRISE /S" + srvnameUD + " /N" + usernameUD + " /P" + userpassUD + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /C\"/EP" + CurDir + "\\ЗапускCOGNOSУД.epf-ExitOnError\"",2,1);

OpenLogFile();
LogWrite("Start UD AP Processing");
CloseLogFile();
WshShell.Run(path_1cv8UD + " ENTERPRISE /S" + srvnameUD + " /N" + usernameUD + " /P" + userpassUD + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /C\"/EP" + CurDir + "\\ЗапускВыгрузкаУДАП.epf-ExitOnError\"",2,1);

OpenLogFile();
LogWrite("!!! END OF EPF !!!");
LogWrite("Copying Log to \\\\sirius\\backup$\\1c_archiv\\1cpromku\\LOGs");
CloseLogFile();

try {
	if(WshFSO.FolderExists("\\\\sirius\\backup$\\1c_archiv\\1cpromku\\LOGs"))
		WshFSO.CopyFile(strLogFileName, "\\\\sirius\\backup$\\1c_archiv\\1cpromku\\LOGs\\", 1); // затереть старый, если есть
	else{
		OpenLogFile();
		LogWrite("error: Folder \"\\\\sirius\\backup$\\1c_archiv\\1cpromku\\LOGs\" not found!");
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

WshShell.Run(CurDir + "\\DelFiles.exe \"\\\\sirius\\backup$\\1c_archiv\\1CPROMKU\\DUMPBASE\" 7 dt",2,1);
WshShell.Run(CurDir + "\\DelFiles.exe \"\\\\sirius\\backup$\\1c_archiv\\1CPROMKU\\DUMPDBCFG\" 7 cf",2,1);