var WshShell = WScript.CreateObject("WScript.Shell");
var Command = "TASKKILL /F /FI \"USERNAME eq 1CARCHIV\" /IM 1cv8.exe";

WshShell.Run(Command,0, true);
