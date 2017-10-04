//////////////////////////////////////////////////////
// Копирование файла в библиотеку SharePoint по URL //
// если файл существует, то он перезаписывается	    //
//////////////////////////////////////////////////////

//************************************************************************
// Входные параметры скрипта: --------------------------------------------
  fileName = "\\\\wtc.loc\\root\\Процессы\\Меню служебной столовой\\Меню.xls";  // Источник - путь к копируемому файлу 
  sharepointUrl = "http://portal/SiteCollectionDocuments"; // Приемник - URL библиотеки/списка SharePoint
  sharepointFileName = sharepointUrl + "/menu_stol.xlsx"; // Конечный URL путь, по которому запишется файл

var WshShell = WScript.CreateObject("WScript.Shell");
var WshFSO = WScript.CreateObject("Scripting.FileSystemObject");
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

//************************************************************************  
// ф-ия получения объекта для работы с HTTP ------------------------------ 
function getXmlHttp(){
  var xmlhttp;
  try {
 	xmlhttp = new ActiveXObject("MSXML2.XMLHTTP");
  } catch (e) {
 	try{
		xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
	}catch(E){
		xmlhttp = false;
 	}
  }
 
  if (!xmlhttp && typeof XMLHttpRequest!='undefined'){
	xmlhttp = new XMLHttpRequest();
  }
  return xmlhttp;
}
//------------------------------------------------------------------------

//************************************************************************
// Основной раздал скрипта: ----------------------------------------------
	OpenLogFile();
	LogWrite("----------------------------------------------");
	LogWrite("----- начало копирования файла на портал -----");
	CloseLogFile();

// составим конечный URL путь файла  
   
  // если файл-источник существует
  if (WshFSO.fileexists(fileName)) { 
  
	try {
		folderTEMP = WshFSO.GetSpecialFolder(2); // Каталог временных файлов
		fileNameXLS = folderTEMP+"\\menu_tmp.xls";
		fileNameXLSX = folderTEMP+"\\menu_stol_tmp.xlsx"; //промежуточный файл, созданный из XLS

		WshFSO.CopyFile(fileName, fileNameXLS, true); // overwrite

		OpenLogFile();
		LogWrite("Преобразуем \""+fileNameXLS+"\" в \""+fileNameXLSX+"\"...");
		CloseLogFile();
		
		// сохранить файл XLS как XLSX
		var Excel = new ActiveXObject("Excel.Application");

		OpenLogFile();
		LogWrite("Получили Объект \"Excel.Application\"...");
		CloseLogFile();

		Excel.Visible = false;
		Excel.DisplayAlerts = false;
		
		//WScript.Sleep(3000); // ждём...						

		Excel.WorkBooks.Open(fileNameXLS);

		OpenLogFile();
		LogWrite("Открыли файл \"" + fileNameXLS + "\"");
		LogWrite("Присвоение имени \"menu\" области данных...");
		CloseLogFile();

		xlLastCell = 11; // 
		LastDataCell = Excel.ActiveCell.SpecialCells(xlLastCell);
		DataRange = Excel.Range(Excel.Cells(1,1), LastDataCell);// зададим имя выделенной области
		DataRange.Name = "menu";

		Excel.ActiveWorkBook.SaveAs(fileNameXLSX,51);

		OpenLogFile();
		LogWrite("Сохранили в файл \"" + fileNameXLSX + "\"");
		LogWrite("Закрытие Excel...");
		CloseLogFile();

		Excel.Quit();
	
		if (WshFSO.fileexists(fileNameXLSX)) {

			OpenLogFile();
			LogWrite("Преобразование выполнено, прочитаем бинарные данные...");
			CloseLogFile();
			
				
									
			// прочитаем данные файла в двоичном виде
			var stream = new ActiveXObject("ADODB.Stream");
			stream.type = 1; // Binary mode
			stream.Open();
			
			// 50 попыток прочитать файл
			for(i=0; i<50; i++){
				
				try {
					flagOK = 1;
					stream.LoadFromFile(fileNameXLSX);
					break;
				}catch(e_file){
					// файл ещё занят, подождём 300 миллисекунд
					flagOK = 0;
					WScript.Sleep(300); // ждём...
				}					
			}		
		  
			if(flagOK==1){
				// дождались и прочитали содержимое файла
				
				var FileData = stream.Read(); // прочитаем данные в объект
				stream.Close();

				// подключимся к SharePoint по URL библиотеки/списка  
				var xmlhttp = getXmlHttp();
				xmlhttp.open('PUT', sharepointFileName, false); // синхронный вызов

				OpenLogFile();
				LogWrite("отправим (\"send\") данные на SharePoint...");
				CloseLogFile();

				// отправим данные на SharePoint
				xmlhttp.send(FileData);	// writing file to SharePoint  

				OpenLogFile();
				LogWrite("Файл обновлён на пртале успешно.");
				CloseLogFile();
			}	
		}else{
			OpenLogFile();
			LogWrite("ОШИБКА. не найден преобразованный файл \"" + fileNameXLSX + "\"");
			CloseLogFile();
	
		} 	
	} catch(e) {
		// error
		OpenLogFile();
		LogWrite("ОШИБКА. " + e.name + ":" + e.message + "\n" + e.stack);
		CloseLogFile();

	}
  } else {
	OpenLogFile();
	LogWrite("ОШИБКА. не найден файл-источник \"" + fileName + "\"");
	CloseLogFile();
  }	

OpenLogFile();
LogWrite("КОНЕЦ скрипта -------------------------------------------");
CloseLogFile();
