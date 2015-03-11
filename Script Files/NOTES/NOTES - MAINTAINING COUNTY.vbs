option explicit
DIM beta_agency, Maintaining_County, worker_position, cao_discussed_check, ButtonPressed, Yes_No, worker_sign, CAAD_Note_text, case_number
'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO					'Declares variables to be good to option explicit users
If beta_agency = "" then 			'For scriptwriters only
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
ElseIf beta_agency = True then		'For beta agencies and testers
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Else								'For most users
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
End if
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
If req.Status = 200 Then									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF
'This script is my maintaining county script that loops until signatue and case number are entered.
'This dialog is used as a CAAD note when Anoka will maintain a case for maintaining county purposes.
'The information below is waht I copied from my Dialog Editor

BeginDialog Maintaining_County, 0, 0, 321, 175, "Maintaining County" 'DEFINING MY DIALOG HERE
  DropListBox 95, 35, 75, 15, "SELECT ONE"+chr(9)+"Yes"+chr(9)+"No", Yes_No
  ComboBox 120, 80, 75, 15, "Anoka Luebben 003/001/cs1/01"+chr(9)+"Pam Scheller 003/001/cs1/02"+chr(9)+"Wendy LeVessuer 003/001/cs1/03", worker_position
  CheckBox 5, 110, 105, 15, "Discussed case with CAO", cao_discussed_check
  EditBox 90, 135, 85, 15, worker_sign
  ButtonGroup ButtonPressed
    OkButton 215, 160, 50, 15
    CancelButton 270, 160, 50, 15
  Text 5, 35, 80, 15, "Will Anoka Maintain?"
  Text 5, 55, 195, 20, "Please send file attn:  Child Supporve move-in.  Transfer Case on PRISM to:"
  Text 5, 80, 110, 20, "Select worker position or fill in postion number "
  Text 5, 140, 85, 10, "Sign your case note"
  Text 5, 10, 55, 10, "Case Number "
  EditBox 65, 10, 110, 15, case_number
EndDialog

'NOW I AM COMMENTING HOW TO PUT IT ALL TOGETHER.
'connecting to PRISM
EMConnect ""

'checks to make sure you are logged into PRISM
CALL check_for_PRISM(True) 'CUSTOM FUNCTION CHEKCING TO MAKE SURE I'M IN PRISM
call PRISM_case_number_finder (case_number)'finds prism number

DO 'creates a Do Loop to create a msgbox if the user forgot to enter the PRISM case number and initial.  REFERENCE YOUR DIALOG HERE.  ALWAYS START YOUR DO LOOP AFTER YOU DEFINE YOUR DIALOG
	Dialog Maintaining_County 'DisplayS the dialog you called above
	If case_number = "" THEN MsgBox "Enter PRISM case number" '
	If worker_sign = "" THEN MsgBox "Please initial script"
LOOP UNTIL case_number <> "" and worker_sign <> "" 'this is saying loop until our case number is filled in and the worker signature is filled in. <> "" means case number is not blank


'Calls to PRISM CAAD
CALL navigate_to_PRISM_screen ("CAAD")

PF5  'adds new caad
EMWriteScreen "A", 3, 29 'enters A in CAAD
EMWriteScreen "T0098", 4, 54 'enters caad code
EMSetCursor 16, 4 'goes to body of caad


If Yes_No= "Yes" THEN 'If the user selected that Anoka will maintain, then create caad note text for transfer information 
	caad_note_text=  "Anoka will maintain" 'writes the text in quotes in CAAD
	call write_bullet_and_variable_in_CAAD ("Maintaining County", caad_note_text) 'makes a bullet and word wraps the info in quotes in 
	call write_bullet_and_variable_in_CAAD ("Transfer case on PRISM to", worker_position) 'makes a bullet and word wraps the info in quotes in caad
	IF cao_discussed_check = 1 THEN 'if user selects that the case was discusses with cao add text
		call write_bullet_and_variable_in_CAAD ("Discussed case with CAO", "") 
	END IF 'the script will end if the worker initially selects Yes.  If they select NO - go to ESLEIF

ELSEIf Yes_No= "No" THEN 'If user selected Anoka will not maintain then create appropriate caad note
	call write_bullet_and_variable_in_CAAD ("Anoka will not maintain", caad_note_text) 
END IF' Have to have END IF again because we started a new ELSEIF

call write_variable_in_CAAD ("-" &worker_sign)'writes the worker signature 



