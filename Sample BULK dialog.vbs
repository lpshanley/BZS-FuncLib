'Try having it run the script from GitHub, but replace the text of FUNCTIONS FILE loading with nothing. That way, it doesn't try and find that file path.???

'LOADING ROUTINE FUNCTIONS-------------------------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/Master Functions Library.vbs"
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
If req.Status = 200 Then									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
End If


'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog pull_data_into_Excel_dialog, 0, 0, 306, 155, "Pull data into Excel dialog"
  ButtonGroup ButtonPressed
    PushButton 10, 30, 25, 10, "ACTV", ACTV_button
    PushButton 35, 30, 25, 10, "EOMC", EOMC_button
    PushButton 60, 30, 25, 10, "PND2", PND2_button
    PushButton 85, 30, 25, 10, "REVS", REVS_button
    PushButton 110, 30, 25, 10, "REVW", REVW_button
    PushButton 135, 30, 25, 10, "MFCM", MFCM_button
	PushButton 10, 60, 25, 10, "ARST", ARST_button
    PushButton 10, 80, 80, 10, "LTC-GRH list generator", LTC_GRH_list_generator_button
    PushButton 10, 105, 75, 10, "SWKR list generator", SWKR_list_generator_button
    CancelButton 250, 135, 50, 15
  Text 5, 5, 125, 10, "What area of REPT are you scanning?"
  GroupBox 5, 20, 160, 25, "Case lists"
  GroupBox 5, 50, 295, 80, "Other"
  Text 40, 60, 250, 20, "--- Caseload stats by worker. Includes cash/SNAP/HC/emergency/GRH stats."
  Text 95, 80, 200, 20, "--- Creates a list of FACIs, AREPs, and waiver types assigned to the various cases in a caseload (or group of caseloads)."
  Text 90, 105, 205, 20, "--- Creates a list of SWKRs assigned to the various cases in a caseload (or group of caseloads)."
EndDialog

'VARIABLES TO DECLARE
all_case_numbers_array = " "					'Creating blank variable for the future array
call worker_county_code_determination(worker_county_code, two_digit_county_code)	'Determines worker county code
is_not_blank_excel_string = Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34)	'This is the string required to tell excel to ignore blank cells in a COUNTIFS function


'THE SCRIPT----------------------------------------------------------------------------------------------------

'Shows report scanning dialog, which asks user which report to generate.
dialog pull_data_into_Excel_dialog
If buttonpressed = cancel then stopscript

'Connecting to BlueZone
EMConnect ""

If buttonpressed = ACTV_button then call run_from_GitHub(script_repository & "BULK - REPT-ACTV list.vbs")
If buttonpressed = ARST_button then call run_from_GitHub(script_repository & "BULK - REPT-ARST list.vbs")
If buttonpressed = EOMC_button then call run_from_GitHub(script_repository & "BULK - REPT-EOMC list.vbs")
If buttonpressed = PND2_button then call run_from_GitHub(script_repository & "BULK - REPT-PND2 list.vbs")
If buttonpressed = REVS_button then call run_from_GitHub(script_repository & "BULK - REPT-REVS list.vbs")
If buttonpressed = REVW_button then call run_from_GitHub(script_repository & "BULK - REPT-REVW list.vbs")
If buttonpressed = MFCM_button then call run_from_GitHub(script_repository & "BULK - REPT-MFCM list.vbs")
If buttonpressed = LTC_GRH_list_generator_button then call run_from_GitHub(script_repository & "BULK - LTC-GRH list generator.vbs")
If buttonpressed = SWKR_list_generator_button then call run_from_GitHub(script_repository & "BULK - SWKR list generator.vbs")

'Logging usage stats
script_end_procedure("If you see this, it's because you clicked a button that, for some reason, does not have an outcome in the script. Contact your alpha user to report this bug. Thank you!")
