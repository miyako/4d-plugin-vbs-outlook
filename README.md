# 4d-plugin-vbs-outlook

#### Office automation from external process no longer an option

Since Office 2013/2016, Outlook [automation](https://support.microsoft.com/en-us/help/196776/office-automation-using-visual-c) has become very difficult. Click-To-Run (a.k.a. C2R) deployment means Outlook no longer exposes interfaces such as [``IConverterSession``](https://msdn.microsoft.com/en-us/library/office/ff960231.aspx) to COM. This one was useful to convert MAPI (``msg``) to MIME (``eml``). Microsoft has decided not to exposes interoperability classes in the common namespace but rather insulate them in their virtual namespace ("bubble").

[``RegOpenKeyEx``](https://msdn.microsoft.com/en-us/library/windows/desktop/ms724862(v=vs.85).aspx) only goes as deep as ``HKLM\SOFTWARE\Microsoft\Office`` and does not give access to ``ClickToRun``. It seems the only way to enable automation is to [hack the registry so that it becomes visible to COM](https://blogs.msdn.microsoft.com/stephen_griffin/2014/04/21/outlook-2013-click-to-run-and-com-interfaces/) by editing the registry.

```bat
reg copy HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\Wow6432Node\CLSID\{4E3A7680-B77A-11D0-9DA5-00C04FD65685} HKLM\SOFTWARE\Classes\Wow6432Node\CLSID\{4E3A7680-B77A-11D0-9DA5-00C04FD65685} /s /f
reg copy HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\Wow6432Node\CLSID\{9EADBD1A-447B-4240-A9DD-73FE7C53A981} HKLM\SOFTWARE\Classes\Wow6432Node\CLSID\{9EADBD1A-447B-4240-A9DD-73FE7C53A981} /s /f
```

**Note**: Must run as Administrator. Both lines (``IConverterSession`` and ``IMimeMessage``) are necessary. Note that the ``15.0`` path component mentioned in the blog post is missing in this example (tested with Office 2016 C2R). This technique is also necessary to make [mfcmapi](https://github.com/stephenegriffin/mfcmapi) work with Office Click-To-Run.

#### Failed attempts

``CoCreateInstance`` crash.  

``LoadLibraryEx`` module not found (no crash) with ``DONT_RESOLVE_DLL_REFERENCES``. Otherwise crash.  

``CoLoadLibrary`` crash (unless used after ``LoadLibraryEx``)   

#### In short

One can no longer depend on Outlook (``IConverterSession``) to export messages in ``eml`` format from an external process.

### VBA to export selected messages

```vba
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
	Else
		exportFolderPath = objArgs(0)
	End if

	If Not objFSO.FolderExists(exportFolderPath) Then
		exportFolderPath = objFSO.CreateFolder(exportFolderPath).Path
	End If

	Set objSelection = objOutlook.ActiveExplorer().Selection

	For i = 1 To objSelection.Count
		Set selObject = objSelection.Item(i)
		exportPath = exportFolderPath & i & ".msg"
		'WScript.StdOut.Write selObject.Body
		selObject.SaveAs exportPath, 9 'olMSGUnicode
		'https://msdn.microsoft.com/en-us/VBA/Outlook-VBA/articles/olsaveastype-enumeration-outlook
	Next

	Set objSelection = Nothing
	Set objArgs = Nothing
	Set objShell = Nothing
	Set objFSO = Nothing
	Set objOutlook = Nothing

End If
```

[MailItem.SaveAs](https://msdn.microsoft.com/en-us/vba/outlook-vba/articles/mailitem-saveas-method-outlook)
