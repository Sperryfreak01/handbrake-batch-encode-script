'################################################################################
'#   Auto HandbrakeCLI Script for Windows XP/Vista/7							#
'#   Copyright (C) 2009-2010  Matt Lovett <mattlovett(at)mattlovett.com>		#
'#   																			#
'#   Based On:																	#
'#	 Auto HandbrakeCLI Script													#
'#   Copyright (C) 2009-2010  Curtis Lee Bolin <curtlee2002(at)gmail.com>		#
'#   																			#
'#	 																			#
'#   This program is free software: you can redistribute it and/or modify		#
'#   it under the terms of the GNU General Public License as published by		#
'#   the Free Software Foundation, either version 3 of the License, or			#
'#   (at your option) any later version.										#
'#																				#
'#   This program is distributed in the hope that it will be useful,			#	
'#   but WITHOUT ANY WARRANTY; without even the implied warranty of				#
'#   MERCHANTaudioBitRateILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the	#
'#   GNU General Public License for more details.								#
'#																				#
'#   You should have received a copy of the GNU General Public License			#
'#   along with this program.  If not, see <http://www.gnu.org/licenses/>.		#
'################################################################################

'================================================================================
'								User Modifible Settings
'================================================================================

'****************************** File Extensions *********************************
'The flowing extensions will be detected and transcoded with HandBrake
Filetype1 = "avi"
Filetype2 = "flv"
Filetype3 = "iso"
Filetype4 = "mkv"
Filetype5 = "mp4"
Filetype6 = "mpeg"
Filetype7 = "mpg"
Filetype8 = "wmv"
Filetype9 = "VOB"
Filetype10 = "zzz"

'****************************** Encoder Settings *********************************
'Location of HandBrake exe files
	strHBlocation 		= "C:\handbrake\HandBrakeCLI.exe"
'Location of MPlayer exe files
	strMPLocation		= "C:\mplayer\mplayer.exe"

	
'conatiner for new video files
	strContainerType	= "mkv"				
'Options that control quality settings of transcoded video
	strVideoSettings	= " --encoder x264 --two-pass --turbo --vb 768 --decomb --loose-anamorphic"
'Video codec settings used when --encoder = x264  uncomment only one setting
	strX264Settings 	= " --x264opts b-adapt=2:rc-lookahead=50"
	'strX264Settings 	= " --x264opts subq=6:partitions=all:8x8dct:me=umh:frameref=5:bframes=3:b-pyramid=1:weightb=1"
	

'Audio settings used when the source files doesn't have AC3 audio
	strNonAACAudio 	= " --audio 1 --aencoder faac --ab 128 --mixdown dpl2 --arate 48 --drc 2.0"
'Audio Settings used when the source files have AC3 audio, default is passthrough
	strAACAudio		= " --audio 1 --aencoder ac3"
	

'Other HandBrake Settings 
strSubtitleSettings = " --native-language eng --subtitle-forced scan --subtitle scan"
strOtherSettings 	= " --markers "
'The folder below will be created and encoded video will be placed there
strOutputFolder 	= "Encoded Files"
strLogFileName		= "Encoder Log.txt"


'================================================================================
'							Modify Below at Your Own Risk
'================================================================================
Const ForAppending = 8
Const ForReading = 1
dim CurPath
dim FileCount			'number of files that will be encoded
dim strCLIcommands		'String  containing the location of HB, CLI paramaters, 
dim Movie2Encode	
dim strAC3 
FileCount 	= 0
CLIcommands = "Error no files to encode"
strAC3 		= "False"	


'Set objFSOScriptPath = CreateObject("Scripting.FileSystemObject")	
CurPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".") & "\" 


'Main routine

DependChk 			'check if the files needed to run script are present in spec'd location
CreateFolder (strOutputFolder)	'creates a folder for encoded files

Set objFSOEncode = CreateObject("Scripting.FileSystemObject")	'connects to system object to read text files
Set objFile = objFSOEncode.OpenTextFile(".\FileList.txt",1)		'opens the list of files with correct extensions

	Do Until objFile.AtEndOfStream				'for each file in the filelist.txt					
		PathofMovie = objFile.ReadLine			'read in the location of the file
		Encode PathofMovie, strHBlocation		'encode the file 
		Logger strCLIcommands, PathofMovie		'log the action taken by the encoder
	Loop										'do it untill all the files have been encoded
	objFile.Close								'closes the system object used to read the filelist
	Wscript.quit								'close the script
	

sub DependChk
	Set objFSOSearch = CreateObject("Scripting.FileSystemObject")	'System calls for accessing file system
		If objFSOSearch.FileExists(strHBlocation) Then				'Looks for handBrake files in location specified in the user settings
			If objFSOSearch.FileExists(strMPLocation) Then			'If Handbrake found look for MPlayer
				FindFiles (".\")									'gather list of files in script directory with correct extensions, if all dependacies located
				If Not FileCount = 0 Then							'Tell the user how many files found and prompt them to continue
					Wscript.Echo "Found " & FileCount & " Video Files" & vbcrlf & " Press OK to begin encoding."
				Else	
					Wscript.Echo "No compatabile files found"		'If no files found tell the user and exit
					Wscript.quit
				End If
			Else													'If MPlayer not found ask user to install it and exit
				Wscript.Echo "Cannot locate Mplayer please download and install to " & strMPLocation
				Wscript.quit
			End If
		Else														'If HandBrake not found prompt user to locate it		
			Wscript.Echo "Cannot locate HandBrake please download and install to " & strHBlocation
			Wscript.quit
		End If
	End Sub

Sub Encode (Movie2Encode, HBPath)
	Set WshShellEncode 	= WScript.CreateObject("WScript.Shell")		'system object to interact with file system
	intTitleLength 		= Len(Movie2Encode) -4						'length of files without exstension	
	strOutputName 		= Left(Movie2Encode, intTitleLength)		'cleanse the extension from the filename
	AC3Check (Movie2Encode)											'detect AC3 audio in the video file
	
	If strAC3 = "True" Then						'If AC3 audio present then generate settings with AC3 audio 
		strCLISettings 	= strVideoSettings & strX264Settings & strAACAudio & strSubtitleSettings & strOtherSettings & " --format " & strContainerType
		strCLIInput 	= " --input " & chr(34) & CurPath & Movie2Encode & chr(34)
		strCLIOuput 	= " --output " & chr(34) & CurPath  & strOutputFolder  & "\" & strOutputName & "." & strContainerType & chr(34)
		strCLICommands	= HBPath  & strCLISettings & strCLIInput & strCLIOuput & strCLICommands
	Else										'If AC3 audio present then generate settings with nonAC3 settings
		strCLISettings 	= strVideoSettings & strX264Settings & strNonAACAudio & strSubtitleSettings & strOtherSettings & " --format " & strContainerType
		strCLIInput 	= " --input " & chr(34) & CurPath & Movie2Encode & chr(34)
		strCLIOuput 	= " --output " & chr(34) & CurPath  & strOutputFolder  & "\" & strOutputName & "." & strContainerType & chr(34)
		strCLICommands	= HBPath  & strCLISettings & strCLIInput & strCLIOuput & strCLICommands
	End If
	
	WshShellEncode.Run strCLICommands, 1, true						'Encodes file with settings gennerated above
		End Sub

Sub Logger (LogEntry, FileEncoded)										'Logs encoder actions
	Set objFSOLogger = CreateObject("Scripting.FileSystemObject")		'System calls to interact with file system																
		If objFSOLogger.FileExists( ".\"  & strOutputFolder  & "\" & strLogFileName  ) Then		'Checks if log file exsits if yes it apends data to file
			Set objFSOCreateLog = CreateObject("Scripting.FileSystemObject")	'Sets log file for apending
			Set objLogFile = objFSOCreateLog.OpenTextFile(".\"  & strOutputFolder  & "\" & strLogFileName , ForAppending)
			objLogFile.WriteLine now 									'print the time the file was encoded
			objLogFile.WriteLine "     Encoded " & FileEncoded			'print the file that was encoded
			objLogFile.WriteLine "     With the following parameters:" 	'buffer text
			objLogfile.WriteLine "     " & LogEntry						'prints the raw paramaters that were passed to handbrake
			objLogFile.Close
			
		Else 
			Set objFSOCreateLog = CreateObject("Scripting.FileSystemObject")	'If text file not present when checked above then create and add data
			Set objLogFile = objFSOCreateLog.CreateTextFile(".\"  & strOutputFolder  & "\" & strLogFileName ) 'creates teh Encoding Log.txt file in same foler as script
			objLogFile.WriteLine now 									'print the time the file was encoded
			objLogFile.WriteLine "     Encoded " & FileEncoded			'print the file that was encoded
			objLogFile.WriteLine "     With the following parameters:" 	'buffer text
			objLogfile.WriteLine "     " & LogEntry						'prints the raw paramaters that were passed to handbrake
			objLogFile.Close
		End If
	
	strCLISettings 		= ""								'clears the string for the next round	
	strCLIInput			= ""								'clears the string for the next round	
	strCLIOuput			= ""								'clears the string for the next round	
	strCLICommands		= ""								'clears the string for the next round	
	End Sub
	
Sub AC3Check (File2Chk)										'Checks for AC3 audio
	strAC3 = ""												'clears the AC3 flag for each iteration
	Set objShell = CreateObject("WScript.Shell")			'system object for file system access
	Set objWshScriptExec = objShell.Exec("C:\mplayer\mplayer.exe -vo null -ao null -frames 0 -identify " & File2Chk)	
	Set objStdOut = objWshScriptExec.StdOut					'Capture the text on the screen 
	
	Do Until objStdOut.AtEndOfStream Or strAC3 = "True"		'check each line cap'd from screen untill done or AC3 found
		strLine = objStdOut.ReadLine
		If strLine = "ID_AUDIO_CODEC=ffac3" Then			'the String indicating AC3 codec being used
			strAC3 = "True"
		Else
			strAC3 = "False"
		End If
	Loop													'repeat until intital conditions met
	End Sub

sub FindFiles (sFolder)
	On Error Resume Next								
	Dim fso, folder, files, NewsFile, strext', sFolder	'declare local variables
  
	Set fso = CreateObject("Scripting.FileSystemObject") 'open file system objects
	'sFolder = Wscript.Arguments.Item(0)					
	'	If sFolder = "" Then							
	'		Wscript.Echo "No Folder parameter was passed"
	'		Wscript.Quit
	'	End If
		Set NewFile = fso.CreateTextFile(".\FileList.txt", True)	'open file system object for writing to a text file
		Set folder = fso.GetFolder(sFolder)				'Open file system obj to read inside a folder
		Set files = folder.Files						'for each loop files equals 1 file in the folder
		FileCount = 0
	
	For each folderIdx In files						'
		strext = Right(folderIdx.Name,3)			'cleanse file name of all but extension
			If strext = Filetype1 Then				'check if the extension matches what the user entered
				NewFile.WriteLine(folderIdx.Name)	'If so then same the file name to the text file
				FileCount = FileCount + 1			'Increase the counter containing the # of files	
			ElseIf strext = Filetype2 Then
				NewFile.WriteLine(folderIdx.Name)	'Repeat for each file extension the user listed
				FileCount = FileCount + 1
			ElseIf strext = Filetype3 Then
				NewFile.WriteLine(folderIdx.Name)
				FileCount = FileCount + 1
			ElseIf strext = Filetype4 Then
				NewFile.WriteLine(folderIdx.Name)
				FileCount = FileCount + 1
			ElseIf strext = Filetype5 Then
				NewFile.WriteLine(folderIdx.Name)
				FileCount = FileCount + 1
			ElseIf strext = Filetype6 Then
				NewFile.WriteLine(folderIdx.Name)
				FileCount = FileCount + 1
			ElseIf ostrext = Filetype7 Then
				NewFile.WriteLine(folderIdx.Name)
				FileCount = FileCount + 1
			ElseIf strext = Filetype8 Then
				NewFile.WriteLine(folderIdx.Name)
				FileCount = FileCount + 1
			ElseIf strext = Filetype9 Then
				NewFile.WriteLine(folderIdx.Name)
				FileCount = FileCount + 1
			ElseIf strext = Filetype10 Then
				NewFile.WriteLine(folderIdx.Name)
				FileCount = FileCount + 1
		End If
	
	Next 
	NewFile.Close									'Close the text file
	End Sub
			

sub CreateFolder (MakeFolder)		
	Set objFSOFolder = CreateObject("Scripting.FileSystemObject")	'System calls for accessing file system
		
	If  objFSOFolder.FolderExists(MakeFolder) Then	'Check if the folder exsists alread
													'If it does do nothing
		Else 										
			objFSOFolder.CreateFolder(MakeFolder)	'If not create the folder
		
		End IF
	
	End Sub 
 