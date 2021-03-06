var WshShell = WScript.CreateObject("WScript.Shell");
var WshFSO = WScript.CreateObject("Scripting.FileSystemObject");
var WshNetwork = WScript.CreateObject("WScript.Network");

CurDir = WshFSO.GetAbsolutePathName(WshShell.CurrentDirectory);

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

srvnameBF = "kappa\\wtc_work_backoffice";
usernameBF = "Exchange";
userpassBF = "12erghm9";
path_1cv8BF = "\"C:\\Program Files (x86)\\1cv8\\8.3.9.2170\\bin\\1cv8.exe\"";

srvnameKU = "kappa\\wtc_work_ku";
usernameKU = "БИТ_ДОлгов";
userpassKU = "Йцукенг8";
path_1cv8KU = "\"C:\\Program Files (x86)\\1cv8\\8.3.9.2170\\bin\\1cv8.exe\"";


OpenLogFile();
LogWrite("");
LogWrite("***** START *****");
LogWrite("---------------------------------------------------------");


LogWrite("START Unloading data from KU to BF");
CloseLogFile();
WshShell.Run(path_1cv8KU + " ENTERPRISE /S" + srvnameKU + " /N" + usernameKU + " /P" + userpassKU + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /Execute\"" + CurDir + "\\8.1\\KU_UnloadtoBF.epf\"",2,1);
OpenLogFile();
LogWrite("END   Unloading data from KU to BF");

LogWrite("START Loading data from KU to BF");
CloseLogFile();

WshShell.Run(path_1cv8BF + " ENTERPRISE /S" + srvnameBF + " /N" + usernameBF + " /P" + userpassBF + " /AU- /Out " + strLogFileName  + " -NoTruncate /RunModeOrdinaryApplication /DisableStartupMessages /C \"ОтключитьЛогикуНачалаРаботыСистемы,ОтключитьЛогикуНачалаРаботыСистемыЦМТ\" /Execute\"" + CurDir + "\\8.3\\BF_LoadFromKU.epf\"",2,1); 

OpenLogFile();
LogWrite("END   Loading data from KU to BF");

LogWrite("START Unloading from BF to KU");
CloseLogFile();

WshShell.Run(path_1cv8BF + " ENTERPRISE /S" + srvnameBF + " /N" + usernameBF + " /P" + userpassBF + " /AU- /Out " + strLogFileName + " -NoTruncate /RunModeOrdinaryApplication /DisableStartupMessages /C \"ОтключитьЛогикуНачалаРаботыСистемы,ОтключитьЛогикуНачалаРаботыСистемыЦМТ\" /Execute\"" + CurDir + "\\8.3\\BF_UnloadToKU.epf\"",2,1);

OpenLogFile();
LogWrite("END   Unloading from BF to KU");

LogWrite("START Loading data from BF to KU");
CloseLogFile();
WshShell.Run(path_1cv8KU + " ENTERPRISE /S" + srvnameKU + " /N" + usernameKU + " /P" + userpassKU + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /Execute\"" + CurDir + "\\8.1\\KU_LoadFromBF.epf\"",2,1);
OpenLogFile();
LogWrite("END   Loading data from BF to KU");


LogWrite("***** The END *****");
CloseLogFile();
