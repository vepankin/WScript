var WshShell = WScript.CreateObject("WScript.Shell");
var WshFSO = WScript.CreateObject("Scripting.FileSystemObject");
CurDir = WshFSO.GetAbsolutePathName(WshShell.CurrentDirectory);

var LogFile;

var startDate = new Date();
startDate.setHours(startDate.getHours()+12);

var dt = new Date();
var MsgDate = String(dt.getYear()) + "-" +( (dt.getMonth() > 8) ? String(dt.getMonth()+1) : ("0" + String(dt.getMonth()+1)) ) + "-" + ((dt.getDate() > 9) ? String(dt.getDate()) : ("0"+String(dt.getDate())) );

var strLogFileName = CurDir + "\\LOGs\\MsgVal_" + MsgDate +".txt";	

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

srvnameUD = "epsilon\\wtc_work_81";
usernameUD = "Администратор";
userpassUD = "admin1csysx";
path_1cv8UD = "\"C:\\Program Files (x86)\\1cv81\\bin\\1cv8.exe\"";

OpenLogFile();
LogWrite("");
LogWrite("***** START *****");
LogWrite("---------------------------------------------------------");

OpenLogFile();
LogWrite("Start UD Currency Load...");
CloseLogFile();
WshShell.Run(path_1cv8UD + " ENTERPRISE /S" + srvnameUD + " /N" + usernameUD + " /P" + userpassUD + " /AU- /Out" + strLogFileName + " -NoTruncate /DisableStartupMessages /UseHWLicenses /C\"/EP" + CurDir + "\\ЗапускЗагрузкаВалютУД.epf-ExitOnError\"",2,1);

OpenLogFile();
LogWrite("***** The END *****");
CloseLogFile();