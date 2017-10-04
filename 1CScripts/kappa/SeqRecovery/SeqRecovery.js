var WshShell = WScript.CreateObject("WScript.Shell");
var WshFSO = WScript.CreateObject("Scripting.FileSystemObject");
var WshNetwork = WScript.CreateObject("WScript.Network");

// завершить зависшие процессы
var Command = "TASKKILL /F /FI \"USERNAME eq SYSTEM\" /IM 1cv8.exe";
WshShell.Run(Command,0, true);


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
usernameBF = "SeqRecovery";
userpassBF = "821egrhm";
path_1cv8BF = "\"C:\\Program Files (x86)\\1cv8\\8.3.9.2170\\bin\\1cv8.exe\"";


OpenLogFile();
LogWrite("---------------------------------------------------------");
LogWrite("START Sequence Recovery");
CloseLogFile();

WshShell.Run(path_1cv8BF + " ENTERPRISE /S" + srvnameBF + " /N" + usernameBF + " /P" + userpassBF + " /WA- /AU- /Out " + strLogFileName + " -NoTruncate /RunModeOrdinaryApplication /DisableStartupMessages /C \"ОтключитьЛогикуНачалаРаботыСистемы,ОтключитьЛогикуНачалаРаботыСистемыЦМТ\" /Execute\"" + CurDir + "\\BF_SeqRecoveryEx.epf\"",2,1);

OpenLogFile();
LogWrite("END Sequence Recovery");
CloseLogFile();
