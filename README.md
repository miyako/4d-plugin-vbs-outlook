# 4d-plugin-vbs-outlook

#### Office automation from external process no longer an option

Since Office 2013/2016, Outlook [automation](https://support.microsoft.com/en-us/help/196776/office-automation-using-visual-c) has become very difficult. Click-To-Run (a.k.a. C2R) deployment means Outlook no longer exposes interfaces such as [``IConverterSession``](https://msdn.microsoft.com/en-us/library/office/ff960231.aspx) to COM. This one was useful to convert MAPI (``msg``) to MIME (``eml``). Microsoft has decided not to exposes interoperability classes in the common namespace but rather insulate them in their virtual namespace ("bubble").

[``RegOpenKeyEx``](https://msdn.microsoft.com/en-us/library/windows/desktop/ms724862(v=vs.85).aspx) only goes as deep as ``HKLM\SOFTWARE\Microsoft\Office`` and does not give access to ``ClickToRun``. It seems the only way to enable automation is to [hack the registry so that it becomes visible to COM](https://blogs.msdn.microsoft.com/stephen_griffin/2014/04/21/outlook-2013-click-to-run-and-com-interfaces/).

```bat
reg copy HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\Wow6432Node\CLSID\{4E3A7680-B77A-11D0-9DA5-00C04FD65685} HKLM\SOFTWARE\Classes\Wow6432Node\CLSID\{4E3A7680-B77A-11D0-9DA5-00C04FD65685} /s /f
reg copy HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\Wow6432Node\CLSID\{9EADBD1A-447B-4240-A9DD-73FE7C53A981} HKLM\SOFTWARE\Classes\Wow6432Node\CLSID\{9EADBD1A-447B-4240-A9DD-73FE7C53A981} /s /f
```

**Note**: Must run as Administrator. Both lines (``IConverterSession`` and ``IMimeMessage``) are necessary. Note that the ``15.0`` path component mentioned in the blog post is missing in this example (tested with Office 2016 C2R). This technique is also necessary to make [mfcmapi](https://github.com/stephenegriffin/mfcmapi) work with Office Click-To-Run.

#### Failed attempts

``CoCreateInstance`` crash.  

``LoadLibraryEx`` module not found (no crash) with ``DONT_RESOLVE_DLL_REFERENCES``. Otherwise crash.  

``CoLoadLibrary`` crash (unless used after ``LoadLibraryEx``)   

#### Alternative methods to convert MAPI to MIME

To parse ``msg``files  without Outlook we could look into [libgsf](https://github.com/GNOME/libgsf) or [COM](https://msdn.microsoft.com/en-us/library/aa380369%28VS.85%29.aspx) based on the specification of 
[Outlook Item File Format](https://msdn.microsoft.com/en-us/library/cc463912%28v=exchg.80%29.aspx?f=255&MSPPError=-2147217396). Once contents are retrieved from the structured file, RTF needs to be converted to HTML. There are numerous libraries designed to do this in [Python](https://github.com/JoshData/convert-outlook-msg-file), [Perl](https://github.com/craig552uk/MSG-Convert) [Perl](https://github.com/mvz/msgconvert), [ruby](https://github.com/aquasync/ruby-msg), [C#](https://github.com/Sicos1977/MSGReader), but they all seem to be quite limited in feature (extract plain text only, for instance).

####  MAPI to MIME is probably the wrong way

As discussed [here](https://blogs.msdn.microsoft.com/stephen_griffin/2008/01/08/no-msg-for-you/), there is a fundamental problem in that ``msg`` is  a faithful copy of the original email and it **should not be used as an archive**. However way Outlook exports a message, they do not represent the original email. Since critical pieces of information are missing in ``msg`` it is simply not possible to "convert" ``msg`` to ``eml``.

Since ``msg`` is a reduced and altered representation of the original email, there is no real incentive in trying to convert ``msg`` to ``eml``. We might as well export from Outlook ([MailItem.SaveAs](https://msdn.microsoft.com/en-us/vba/outlook-vba/articles/mailitem-saveas-method-outlook)) in ``MHT`` format. The content is of course not the same as the original, but no less so than ``msg``, it will contain all the attachments and [``MIME``](https://en.wikipedia.org/wiki/MIME) is easier to parse than [``CFBF``](https://en.wikipedia.org/wiki/Compound_File_Binary_Format).

#### VBA to export selected messages

```vba
CRLF = Chr(13) & Chr(10)
On Error Resume Next
	Set objOutlook = GetObject("", "Outlook.Application") 'empty string must be explicit
On Error GoTo 0

If Err.Number = 0 Then

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objShell = WScript.CreateObject("Shell.Application")
	Set objArgs = WScript.Arguments

	If objArgs.Count = 0 Then
		exportFolderPath = objShell.Namespace(0).Self.Path 'Desktop
		'https://msdn.microsoft.com/en-us/library/windows/desktop/bb774096.aspx
		WScript.StdOut.Write "no path specified, default to desktop" & CRLF
	Else
		exportFolderPath = objArgs(0)
	End if

	If Not objFSO.FolderExists(exportFolderPath) Then
		exportFolderPath = objFSO.CreateFolder(exportFolderPath).Path
		WScript.StdOut.Write "creating folder " & exportFolderPath & CRLF
	End If

	If Not Right(Trim(exportFolderPath), 1) = "\" Then
		exportFolderPath = exportFolderPath & "\"
	End If

	WScript.StdOut.Write "export to " & exportFolderPath & CRLF

	Set objSelection = objOutlook.ActiveExplorer().Selection

	For i = 1 To objSelection.Count
		Set selObject = objSelection.Item(i)
		exportPath = exportFolderPath & i & ".mht"
		'WScript.StdOut.Write selObject.Body
		On Error Resume Next
			selObject.SaveAs exportPath, 10 'olMHTML
			'https://msdn.microsoft.com/en-us/VBA/Outlook-VBA/articles/olsaveastype-enumeration-outlook
		On Error GoTo 0
		WScript.StdOut.Write "creating file " & exportPath & CRLF
		WScript.StdOut.Write "result code " & Err.Number & CRLF
	Next

	Set objSelection = Nothing
	Set objArgs = Nothing
	Set objShell = Nothing
	Set objFSO = Nothing
	Set objOutlook = Nothing

End If
```

#### Get the list of messages dropped from Outlook in 4D

The clipboard data type [``CFSTR_FILEDESCRIPTORW``](https://msdn.microsoft.com/en-us/library/windows/desktop/bb776902(v=vs.85).aspx) (``FileGroupDescriptorW``) is available during the ``On Drop`` form event or the first ``On Drag Over`` event in 4D.

Its structure is known ([``FILEGROUPDESCRIPTOR``](https://msdn.microsoft.com/en-us/library/windows/desktop/bb773290(v=vs.85).aspx), [``FILEDESCRIPTOR``](https://msdn.microsoft.com/en-us/library/windows/desktop/bb773288(v=vs.85).aspx)) so we can parse it with regular 4D commands. 

```
C_BLOB($1)
C_OBJECT($0;$FileGroupDescriptor)

If (Count parameters#0)
	
	$sizeof_INPUT:=BLOB size($1)
	
	$sizeof_UINT:=4
	$sizeof_DWORD:=8
	$sizeof_CLSID:=16
	$sizeof_FILETIME:=8  //DWORD*2
	$sizeof_SIZEL:=8  //LONG*2
	$sizeof_POINTL:=8  //LONG*2
	$sizeof_TCHAR_MAX_PATH:=520  //260*sizeof(wchar_t)
	
	  //size test #1
	If (($sizeof_INPUT%592)=$sizeof_UINT)
		
		C_LONGINT($pos)
		ARRAY OBJECT($FileDescriptors;0)
		
		  //UINT           cItems;
		$cItems:=BLOB to longint($1;PC byte ordering;$pos)
		$sizeof_fgd:=($cItems+1)*592
		
		  //size test #2
		If (($sizeof_INPUT-$pos)=$sizeof_fgd)
			
			C_BLOB($clsid)
			C_BLOB($sizel;$pointl)
			C_BLOB($ftCreationTime;$ftLastAccessTime;$ftLastWriteTime)
			C_BLOB($fileName)
			
			For ($i;1;$cItems)
				
				  //DWORD    dwFlags;
				$dwFlags:=BLOB to longint($1;PC byte ordering;$pos)
				
				  //CLSID    clsid;
				COPY BLOB($1;$clsid;$pos;0;$sizeof_CLSID)
				$pos:=$pos+$sizeof_CLSID
				
				  //SIZEL    sizel;
				COPY BLOB($1;$sizel;$pos;0;$sizeof_SIZEL)
				$pos:=$pos+$sizeof_SIZEL
				
				  //POINTL   pointl;
				COPY BLOB($1;$pointl;$pos;0;$sizeof_POINTL)
				$pos:=$pos+$sizeof_POINTL
				
				  //DWORD    dwFileAttributes;
				$dwFileAttributes:=BLOB to longint($1;PC byte ordering;$pos)
				
				  //FILETIME ftCreationTime;
				COPY BLOB($1;$ftCreationTime;$pos;0;$sizeof_FILETIME)
				$pos:=$pos+$sizeof_FILETIME
				
				  //FILETIME ftLastAccessTime;
				COPY BLOB($1;$ftLastAccessTime;$pos;0;$sizeof_FILETIME)
				$pos:=$pos+$sizeof_FILETIME
				
				  //FILETIME ftLastWriteTime;
				COPY BLOB($1;$ftLastWriteTime;$pos;0;$sizeof_FILETIME)
				$pos:=$pos+$sizeof_FILETIME
				
				  //DWORD    nFileSizeHigh;
				$nFileSizeHigh:=BLOB to longint($1;PC byte ordering;$pos)
				
				  //DWORD    nFileSizeLow;
				$nFileSizeLow:=BLOB to longint($1;PC byte ordering;$pos)
				
				  //TCHAR    cFileName[MAX_PATH];
				COPY BLOB($1;$fileName;$pos;0;$sizeof_TCHAR_MAX_PATH)
				  //trim at wcslen (null bytes survive in object)
				$cFileName:=Convert to text($fileName;"utf-16le")
				$cFileName:=Substring($cFileName;1;Position(Char(0);$cFileName;*)-1)
				
				$pos:=$pos+$sizeof_TCHAR_MAX_PATH
				
				C_OBJECT($fgd)
				
				OB SET($fgd;"cFileName";$cFileName)
				APPEND TO ARRAY($FileDescriptors;$fgd)
				CLEAR VARIABLE($fgd)
				
			End for 
			
		End if 
		
		OB SET($FileGroupDescriptor;"cItems";$cItems)
		OB SET ARRAY($FileGroupDescriptor;"fgt";$FileDescriptors)
		
	End if 
End if 

$0:=$FileGroupDescriptor
```
