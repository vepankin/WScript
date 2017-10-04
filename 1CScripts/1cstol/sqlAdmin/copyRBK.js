var WshShell = WScript.CreateObject("WScript.Shell");
var WshFSO = WScript.CreateObject("Scripting.FileSystemObject");
CurDir = WshFSO.GetParentFolderName(WScript.ScriptFullName);


var LogFile;
var dt = new Date();
var MsgDate = String(dt.getYear()) + "-" +( (dt.getMonth() > 9) ? String(dt.getMonth()+1) : ("0" + String(dt.getMonth()+1)) ) + "-" + ((dt.getDate() > 9) ? String(dt.getDate()) : ("0"+String(dt.getDate())) );

var strLogFileName = CurDir + "\\LOGs\\MsgCopyRBK_" + MsgDate +".txt";	

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

//*************************************************************************************************************

function MyCopyFolder(aPath,aDestPath){

var f, fc, s, ff;

   f = WshFSO.GetFolder(aPath);
   fc = new Enumerator(f.files);
   
   s = "";   

   for (; !fc.atEnd(); fc.moveNext()){
      s = aPath + "\\" + fc.item().Name;

      try {
         WshFSO.CopyFile(s, aDestPath + "\\", 1); // затереть старый, если есть
      }catch(e){
		// error
		OpenLogFile();
		LogWrite("Error on Copy file "+ s + " " + e.message);
		CloseLogFile();

      }	
   }


   ff = new Enumerator(f.SubFolders);
   
   for (; !ff.atEnd(); ff.moveNext()){
      s = aPath + "\\" + ff.item().Name;

      if(!WshFSO.FolderExists(aDestPath + "\\" + ff.item().Name))
         WshFSO.CreateFolder(aDestPath + "\\" + ff.item().Name);


      MyCopyFolder(s,aDestPath + "\\" + ff.item().Name);
   }
     
   return(0);

} // end function GetLastFileName

//*************************************************************************************************************
OpenLogFile();
LogWrite("Start copyRBK...");
CloseLogFile();

try{
	if((dt.getDate()%2 ) == 0){

	   OpenLogFile();
	   LogWrite("Copying to D:\\RBK_COPY\\1_COPY");
	   CloseLogFile();

	   MyCopyFolder("D:\\RBK",  "D:\\RBK_COPY\\1_COPY");


	   OpenLogFile();
	   LogWrite("Copying to sirius - 1_COPY");
	   CloseLogFile();

	   MyCopyFolder("D:\\RBK",  "\\\\sirius\\backup$\\1C_ARCHIV\\1CSTOL\\STOLRBK\\1_COPY");

	   OpenLogFile();
	   LogWrite("End copying");
	   CloseLogFile();


	}else{

	   OpenLogFile();
	   LogWrite("Copying to D:\\RBK_COPY\\2_COPY");
	   CloseLogFile();

	   MyCopyFolder("D:\\RBK",  "D:\\RBK_COPY\\2_COPY");

	   OpenLogFile();
	   LogWrite("Copying to sirius - 2_COPY");
	   CloseLogFile();

	   MyCopyFolder("D:\\RBK",  "\\\\sirius\\backup$\\1C_ARCHIV\\1CSTOL\\STOLRBK\\2_COPY");

	   OpenLogFile();
	   LogWrite("End copying");
	   CloseLogFile();

	}
}catch(e){
	// error
	OpenLogFile();
	LogWrite("Error copying: "+e.message);
	CloseLogFile();
}

OpenLogFile();
LogWrite("End");
CloseLogFile();

// copy the log to Sirius
WshFSO.CopyFile(strLogFileName, "\\\\sirius\\backup$\\1C_ARCHIV\\1CSTOL\\STOLRBK\\RBKCopyLog.txt", 1); // затереть старый, если есть 