# 4d-plugin-vbs-outlook

Experimental

Since Office 2013/2016, Outlook automation has become very difficult. Click-To-Run (a.k.a. C2R) deployment means Outlook no longer exposes interfaces such as [``IConverterSession``](https://msdn.microsoft.com/en-us/library/office/ff960231.aspx) to COM. This one was useful to convert MAPI (``msg``) to MIME (``eml``). Microsoft has decided not to exposes interoperability classes in the common namespace but rather insulate them in their virtual namespace ("bubble").

[``RegOpenKeyEx``](https://msdn.microsoft.com/en-us/library/windows/desktop/ms724862(v=vs.85).aspx) only goes as deep as ``HKLM\SOFTWARE\Microsoft\Office`` and does not give access to ``ClickToRun``. It seems like the only way to enable automation is to use the [hack to register the class in the namespace visible to COM](https://blogs.msdn.microsoft.com/stephen_griffin/2014/04/21/outlook-2013-click-to-run-and-com-interfaces/) by editing the registry.

```bat
reg copy HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\Wow6432Node\CLSID\{4E3A7680-B77A-11D0-9DA5-00C04FD65685} HKLM\SOFTWARE\Classes\Wow6432Node\CLSID\{4E3A7680-B77A-11D0-9DA5-00C04FD65685} /s /f
```

**Note**: must run as Administrator. Compared to the blog post, the ``15.0`` path component is missing with Office 2016. In any case, code crashes with ``CoCreateInstance`` (same error with ``LoadLibraryEx``, although the latter doesn't crash). It seems the only way to automate Outlook is to use [add-ons](https://msdn.microsoft.com/en-us/library/jj900714.aspx#Automating%20Outlook%20by%20in-process%20vs.%20out-of-process%20Solutions).

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
