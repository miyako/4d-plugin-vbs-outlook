# 4d-plugin-vbs-outlook

Experimental

Since Office 2013/2016, Outlook automation has become very difficult. Click-To-Run (a.k.a. C2R) deployment means Outlook no longer exposes interfaces such as [``IConverterSession``](https://msdn.microsoft.com/en-us/library/office/ff960231.aspx) to COM. This one was useful to convert MAPI (``msg``) to MIME (``eml``). Microsoft has decided not to exposes interoperability classes in the common namespace but rather insulate them in their virtual namespace ("bubble").


There is a [hack to register the class in the com namespace](https://blogs.msdn.microsoft.com/stephen_griffin/2014/04/21/outlook-2013-click-to-run-and-com-interfaces/). However, ``RegOpenKeyEx`` only goes as deep as ``SOFTWARE\Microsoft\Office`` and does not seem to give access to ``ClickToRun`` (only the following sub-keys are visible).

``15.0``  
``16.0``  
``Common``  
``Delivery``  
``Excel``  
``MS Project``  
``Outlook``  
``PowerPoint``  


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
