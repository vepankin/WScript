//var WshShell = WScript.CreateObject("WScript.Shell");
//var WshFSO = WScript.CreateObject("Scripting.FileSystemObject");
var WshShellApp = WScript.CreateObject("Shell.Application");

// останавливаем сервис, пока не остановится
do{
	WshShellApp.ServiceStop("1C:Enterprise 8.1 Server Agent",false);
	WScript.Sleep(2000);

}while(WshShellApp.IsServiceRunning("1C:Enterprise 8.1 Server Agent"));

// запускаем сервис, пока не запустится
do{
	WshShellApp.ServiceStart("1C:Enterprise 8.1 Server Agent",false);
	WScript.Sleep(2000);

}while(!WshShellApp.IsServiceRunning("1C:Enterprise 8.1 Server Agent"));

WScript.Sleep(1000);


