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


'================================================================================
'							Modify Below at Your Own Risk
'================================================================================
CLIcommands = "Error no files to encode"

dim FileCount			'number of files that will be encoded
dim strCLIcommands		'String  containing the location of HB, CLI paramaters, 
dim Movie2Encode	
dim strAC3 

strAC3 = "False"	

'Main routine

'Set objFSOSearch = CreateObject("Scripting.FileSystemObject")			'System calls for accessing file system
'If objFSOSearch.FileExists(strHBlocation) Then							'Looks for handBrake files in location specified in the user settings
		'Wscript.Echo "Found HandBrakeCLI"                				'Alterts user that handbrake has been found
'		FindFiles (".\")												'gather files in script directory with correct extensions, lists them in FileList.txt
	'	Wscript.Echo "Found HandBrakeCLI" & vbcrlf & vbcrlf & "Found " & FileCount & " Video Files" & vbcrlf & " Press OK to begin encoding."
		DependChk
		CreateFolder (strOutputFolder)
		
		Set objFSOEncode = CreateObject("Scripting.FileSystemObject")
		Set objFile = objFSOEncode.OpenTextFile(".\FileList.txt",1)
	
		Do Until objFile.AtEndOfStream
			PathofMovie = objFile.ReadLine
			Encode PathofMovie, strHBlocation
			Logger strCLIcommands, Movie2Encode	
		Loop
		objFile.Close
		Wscript.quit
	
	'Else																'If HandBrake not found prompt user to locate it		
	'	Wscript.Echo "Cannot locate HandBrake please download and install to c:\handbrake"
	'	Wscript.quit
'End If

sub DependChk
	Set objFSOSearch = CreateObject("Scripting.FileSystemObject")			'System calls for accessing file system
	If objFSOSearch.FileExists(strHBlocation) Then							'Looks for handBrake files in location specified in the user settings
		'Wscript.Echo "Found HandBrakeCLI"                				'Alterts user that handbrake has been found
		FindFiles (".\")												'gather files in script directory with correct extensions, lists them in FileList.txt
		Wscript.Echo "Found HandBrakeCLI" & vbcrlf & vbcrlf & "Found " & FileCount & " Video Files" & vbcrlf & " Press OK to begin encoding."
		'CreateFolder (strOutputFolder)
		'
		'Set objFSOEncode = CreateObject("Scripting.FileSystemObject")
		'Set objFile = objFSOEncode.OpenTextFile(".\FileList.txt",1)
	'
		'Do Until objFile.AtEndOfStream
		'	PathofMovie = objFile.ReadLine
		'	Encode PathofMovie, strHBlocation
		'	Logger strCLIcommands, Movie2Encode	
		'Loop
		'objFile.Close
		'Wscript.quit
	
	Else																'If HandBrake not found prompt user to locate it		
		Wscript.Echo "Cannot locate HandBrake please download and install to c:\handbrake"
		Wscript.quit
	End If
	End Sub

Sub Encode (Movie2Encode, HBPath)
	Set WshShellEncode = WScript.CreateObject("WScript.Shell")
	intTitleLength = Len(Movie2Encode) -4
	strOutputName = Left(Movie2Encode, intTitleLength)
	AC3Check (Movie2Encode)
	If strAC3 = "True" Then
		'strCLIcommands = HBPath & strVideoSettings & strX264Settings & strAACAudio & strSubtitleSettings & strOtherSettings & " --format " & strContainerType & " --input " & Movie2Encode & " --output " & ".\" & strOutputFolder & "\" & strOutputName & "." & strContainerType
	wscript.echo strAC3 & vbcrlf & strAACAudio
	Else
		'strCLIcommands = HBPath & strVideoSettings & strX264Settings & strNonAACAudio & strSubtitleSettings & strOtherSettings & " --format " & strContainerType & " --input " & Movie2Encode & " --output " & ".\" & strOutputFolder & "\" & strOutputName & "." & strContainerType
	wscript.echo strAC3 & vbcrlf & strNonAACAudio
	End If
	WshShellEncode.Run strCLIcommands, 1, true								'Begins the encoding proceess using switchs defined in strings at begging of script
	End Sub

Sub Logger (LogEntry, FileEncoded)										'Logging routine saves txt file in the same folder as the script detailing files encoded, 
	Set objFSOLogger = CreateObject("Scripting.FileSystemObject")		'Inports file system calls																'if files exsists apends the file, otherwise creates the log file and writes data
		If objFSOLogger.FileExists(".\" & strOutputFolder & "\" & "Encoding Log.txt") Then				'Checks if log file exsits if yes it apends data to file
			Const ForAppending = 8
			Set objFSOCreateLog = CreateObject("Scripting.FileSystemObject")	'Info below is the information apended to the txt file
			Set objLogFile = objFSOCreateLog.OpenTextFile(".\" & strOutputFolder & "\" & "Encoding Log.txt", ForAppending)
			objLogFile.WriteLine vbcrlf & now & vbcrlf & "     Encoded " & FileEncoded & vbcrlf & "     With the following parameters:" & vbcrlf & "     " & LogEntry
			objLogFile.Close
			
		Else 
			Set objFSOCreateLog = CreateObject("Scripting.FileSystemObject")	'If text file not present when checked above then create and add data
			Set objLogFile = objFSOCreateLog.CreateTextFile(".\" & strOutputFolder & "\" & "Encoding Log.txt") 'creates teh Encoding Log.txt file in same foler as script
			objLogFile.WriteLine vbcrlf & now & vbcrlf & "     Encoded " & FileEncoded & vbcrlf & "     With the following parameters:" & vbcrlf & "     " & LogEntry
			objLogFile.Close
		End If
	End Sub
	
Sub AC3Check (File2Chk)

	Set objShell = CreateObject("WScript.Shell")	
	Set objWshScriptExec = objShell.Exec("C:\mplayer\mplayer.exe -vo null -ao null -frames 0 -identify" & File2Chk)
	Set objStdOut = objWshScriptExec.StdOut
	
	Do Until objStdOut.AtEndOfStream Or strAC3 = "True"
		strLine = objStdOut.ReadLine
		If strLine = "ID_AUDIO_CODEC=ffac3" Then
			strAC3 = "True"
		Else
			strAC3 = "False"
		End If
	Loop
	End Sub

sub FindFiles (sFolder)
	On Error Resume Next
	Dim fso, folder, files, NewsFile, strext', sFolder
  
	Set fso = CreateObject("Scripting.FileSystemObject")
	sFolder = Wscript.Arguments.Item(0)
		If sFolder = "" Then
			Wscript.Echo "No Folder parameter was passed"
			Wscript.Quit
		End If
		Set NewFile = fso.CreateTextFile(".\FileList.txt", True)
		Set folder = fso.GetFolder(sFolder)
		Set files = folder.Files
		FileCount = 0
	
	For each folderIdx In files
		strext = Right(folderIdx.Name,3)
			If strext = Filetype1 Then
				NewFile.WriteLine(folder & "\" & folderIdx.Name)
				FileCount = FileCount + 1
			ElseIf strext = Filetype2 Then
				NewFile.WriteLine(folder & "\"& folderIdx.Name)
				FileCount = FileCount + 1
			ElseIf strext = Filetype3 Then
				NewFile.WriteLine(folder & "\"& folderIdx.Name)
				FileCount = FileCount + 1
			ElseIf strext = Filetype4 Then
				NewFile.WriteLine(folder & "\"& folderIdx.Name)
				FileCount = FileCount + 1
			ElseIf strext = Filetype5 Then
				NewFile.WriteLine(folder & "\"& folderIdx.Name)
				FileCount = FileCount + 1
			ElseIf strext = Filetype6 Then
				NewFile.WriteLine(folder & "\"& folderIdx.Name)
				FileCount = FileCount + 1
			ElseIf ostrext = Filetype7 Then
				NewFile.WriteLine(folder & "\"& folderIdx.Name)
				FileCount = FileCount + 1
			ElseIf strext = Filetype8 Then
				NewFile.WriteLine(folder & "\"& folderIdx.Name)
				FileCount = FileCount + 1
			ElseIf strext = Filetype9 Then
				NewFile.WriteLine(folder & "\"& folderIdx.Name)
				FileCount = FileCount + 1
			ElseIf strext = Filetype10 Then
				NewFile.WriteLine(folder & "\"& folderIdx.Name)
				FileCount = FileCount + 1
		End If
	
	Next 
	NewFile.Close
	End Sub
			

			
		
sub CreateFolder (MakeFolder)		
	'Create FileSystemObject. So we can apply .createFolder method
	Set objFSOFolder = CreateObject("Scripting.FileSystemObject")
	On Error Resume Next
	objFSOFolder.CreateFolder(".\" & MakeFolder)
	End Sub 

 