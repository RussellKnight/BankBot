
'Note that with VBS there is no type declaration like there is with VBA

'Check that we are connected to the Internet and quit if not connected
If pingsilent("4.2.2.2") = false then  Wscript.Quit

'This is the workbook
strPath = "C:\YourPathHere\BankBot 1.01.xlsm" 

'Write the macro name - could try including module name
strMacro = "Module1.Transactions" 'This is the Macro name to execute the API call for transactions

On Error Resume Next
Set xl = GetObject(, "Excel.Application")  'attach to running Excel instance

Set wb = Nothing

'Check all open workbooks to establish whether our workbook is currently open
For Each obj In xl.Workbooks
  If obj.Name = "BankBot 1.01.xlsm" Then  'use obj.FullName for full path
    'We have found the workbook open and thus assign it to the variable
	Set wb = obj
    Exit For
  End If
Next

If wb Is Nothing Then 'If the workbook is not currently open
  'Create an Excel instance and set visibility of the instance
  Set objApp = CreateObject("Excel.Application")
  
  objApp.Visible = True   'Sets the visibility of Excel 
  
  'Open workbook
  Set wbToRun = objApp.Workbooks.Open(strPath) 
  

  'Run macro to execute 
  objApp.Run strMacro     'Run the Macro 
  
  wbToRun.Save
  
  'Wait a second for the save to take place otherwise there is a synching issue with OneDrive
  wscript.sleep 4000
  
  'Close workbook (should result in automatic save)
  wbToRun.Close 
  
  'Close the new instance of Excel that was created (this quit action should not impact other Excel instances that may be open)
  objApp.Quit
Else 'If the file is already open then run the macro from the existing session
  
  xl.run strMacro
  'Decided not to close the workbook as it was open to begin with
  'wb.Close
End If

Wscript.Quit


Function pingsilent (strComputer)
    pingsilent = false
	Dim ReturnCode

    set objShell = WScript.CreateObject("WScript.Shell")
    'set objExec = objShell.Exec("%comspec% /c ping.exe " & strComputer & " -n 1 -w 2000")
    
	ReturnCode = objShell.Run("ping -n 1 -w 2000 " & strComputer,0,True)

   If returncode = 0 then
       pingsilent = true
       exit function
   end if
End function

