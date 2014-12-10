'Try having it run the script from GitHub, but replace the text of FUNCTIONS FILE loading with nothing. That way, it doesn't try and find that file path.???

'LOADING GLOBAL VARIABLES
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\Users\pwvkc45\Desktop\GitHub\BZS-FuncLib\SETTINGS - Global variables.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'LOADING SCRIPT
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/Sample BULK dialog.vbs"
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
If req.Status = 200 Then									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
End If