var WshShell = WScript.CreateObject("WScript.Shell");
var WshFSO = WScript.CreateObject("Scripting.FileSystemObject");
var WshNetwork = WScript.CreateObject("WScript.Network")

var colDrives = WshNetwork.EnumNetworkDrives(); 
if (colDrives.length != 0) { 
		
	for (i = 0; i < colDrives.length; i += 2) { 
		if(colDrives(i)=="T:")WshNetwork.RemoveNetworkDrive("T:",1,1);  
	} 	
}
WshNetwork.MapNetworkDrive("T:","\\\\sirius\\backup$\\1C_ARCHIV\\1CSTOL\\STOLBASE",0,"WTC\\1carchiv","88881111");


//CurDir = WshFSO.GetAbsolutePathName(WshShell.CurrentDirectory);
//CurDir = "D:\SQLADMIN";
CurDir = WshFSO.GetParentFolderName(WScript.ScriptFullName);


var LogFile;
var dt = new Date();
var MsgDate = String(dt.getYear()) + "-" +( (dt.getMonth() > 9) ? String(dt.getMonth()+1) : ("0" + String(dt.getMonth()+1)) ) + "-" + ((dt.getDate() > 9) ? String(dt.getDate()) : ("0"+String(dt.getDate())) );

var strLogFileName = CurDir + "\\LOGs\\DiffBakStol_" + MsgDate +".txt";	

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

//********************************************************************
//********************************************************************
function GetFileExtension(aName){
var i,l,s;
   s = "";
   l = aName.length;

   for(i=2; i<l; i++){
   	if(aName.charAt(l-i)=='.'){
		s = aName.substr(l-i,i);
		break;
	}
   }

   return(s.toUpperCase()); // e.g. ".txt"
} // end function GetFileExtension
//********************************************************************
//********************************************************************
function GetLastFileName(aPath, aExt){
var f, fc, s, vDate;
   f = WshFSO.GetFolder(aPath);
   fc = new Enumerator(f.files);
   s = "";   
   aExt = aExt.toUpperCase();

   if(!fc.atEnd()){

	for (; !fc.atEnd() && (s==""); fc.moveNext()){
		if(GetFileExtension(fc.item().Name)!=aExt) 
			continue;

		vDate = fc.item().DateLastModified;
		s = fc.item().Name;
   	}


	for (; !fc.atEnd(); fc.moveNext()){
		if(GetFileExtension(fc.item().Name)!=aExt) 
			continue;

      		if(vDate < fc.item().DateLastModified){
			s = fc.item().Name;
			vDate = fc.item().DateLastModified;
      		}
   	}
   }else{
	s = ""; // no files
  }
     
   return(s);

} // end function GetLastFileName
//********************************************************************
//********************************************************************



OpenLogFile();
LogWrite("Start copyDiffBAK...");
LogWrite("Delete all BUT FIFTEEN last full copies from Sirius");
CloseLogFile();


try{
	// Delete all BUT FIVE last full copies from sirius
	WshShell.Run(CurDir + "\\DelFiles.exe \"T:\\DIFF\" 15 BAK",2,1);

}catch(e){
	// error
}


// Get the last created file on D:
sLastFileNameBAK1 = GetLastFileName("D:\\SQL_COPY\\DIFF",".BAK");


// Get the last created file on sirius
sLastFileNameBAK2 = GetLastFileName("T:\\DIFF",".BAK");

if(sLastFileNameBAK1 != sLastFileNameBAK2){

	OpenLogFile();
	LogWrite("Copying the last full copy to sirius");
	CloseLogFile();

	// Copy the last full copy to sirius
	try {
		WshFSO.CopyFile("D:\\SQL_COPY\\DIFF\\"+sLastFileNameBAK1, "T:\\DIFF\\", 1); // затереть старый, если есть
	}catch(e){
		// error
		OpenLogFile();
		LogWrite("Error on Copy the last full copy to \\\\sirius "+e.message);
		CloseLogFile();
	}
}


OpenLogFile();
LogWrite("End");
CloseLogFile();
