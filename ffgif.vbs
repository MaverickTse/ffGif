Rem Check if CS or WS is running
Dim isCS
strEngine= UCase(Right(WScript.FullName, 11))
If (strEngine = "CSCRIPT.EXE") Then
isCS = True
Else
isCS =False
End If

Dim objShell, objFS, objStdOut

Set objShell= CreateObject("WScript.Shell")
Set objFS= CreateObject("Scripting.FileSystemObject")

If isCS Then
  Set objStdOut= WScript.StdOut
End If

Dim iColor, iFPS, iMethod, iLoop, strInput, strTransparent
iColor=256
iFPS=24
iMethod=4
iLoop=0
strTransparent=""
Dim cmd1, cmd2, cmd3, strPalettePath, strGifPath, strGifOptPath
Dim oExec
aDitherMethod = Array("bayer","heckbert","floyd_steinberg","sierra2","sierra2_4a")
strScriptFolder= objFS.GetParentFolderName(WScript.ScriptFullName)
strFFPath= """" & strScriptFolder & "\ffmpeg.exe" & """"
strGifsiclePath= """" & strScriptFolder & "\gifsicle.exe" & """"
If (WScript.Arguments.Unnamed.length > 0) Then
  Rem get user input for encode parameters
  strInput= InputBox("Max color count:","Encode parameter", "256")
  If(Not IsEmpty(strInput)) And (Not IsNull(strInput)) Then
  iColor= CInt(strInput)
  Else
  iColor=256
  End If
  strInput= InputBox("Target frame rate","Encode parameter", "24")
  If(Not IsEmpty(strInput)) And (Not IsNull(strInput)) Then
  iFPS= CInt(strInput)
  Else
  iFPS=24
  End If
  strInput= InputBox("Loop count","Encode parameter", "0")
  If(Not IsEmpty(strInput)) And (Not IsNull(strInput)) Then
  iLoop= CInt(strInput)
  Else
  iLoop=0
  End If
  strInput= InputBox("[0]:bayer(ordered)" & vbCrLf & "[1]:heckbert" & vbCrLf & "[2]:floyd_steinberg" & vbCrLf & "[3]:sierra2" & vbCrLf & "[4]:sierra-lite","Dither Method", "4")
  If(Not IsEmpty(strInput)) And (Not IsNull(strInput)) Then
  iMethod= CInt(strInput)
  Else
  iMethod=4
  End If
  Rem Force transparent color
  strInput= InputBox("Set Background Transparency?" & vbCrLf & "For Green BG: #00FF00" & vbCrLf & "For Blue BG: #0000FF" & vbCrLf & "Leave blank to skip this feature" & vbCrLf & "BG transparency will only be applied to" & vbCrLf & "Gifsicle-optimized files","Encode parameter", "")
  If(Len(strInput)=7) Then
  strTransparent= "-B=" & strInput & " -t=" & strInput & " -D=2"
  Else
  strTransparent=""
  End If
  Rem range check for user input
  If (iColor<4) Then iColor=4 End If
  If (iColor>256) Then iColor=256 End If
  If (iFPS<1) Then iFPS=1 End If
  If (iFPS>120) Then iFPS=120 End If
  If (iLoop<0) Then iLoop=0 End If
  If (iLoop>255) Then iLoop=255 End if
  If (iMethod<0) Then iMethod=4 End If
  If (iMethod>4) Then iMethod=4 End If
  Rem loop through each file
  For Each path In WScript.Arguments.Unnamed
  strTargetFolder= objFS.GetParentFolderName(path) & "\"
  strBaseName= objFS.GetBaseName(path)
  strPalettePath= """" & strTargetFolder & strBaseName & "_palette.png" & """"
  strGifPath= """" & strTargetFolder & strBaseName & ".gif" & """"
  strGifOptPath=  """" & strTargetFolder & strBaseName & "_optimized.gif" & """"
  strSrcPath= """" & path & """"
  Rem palettegen commandline
  cmd1= strFFPath & " -y -i " & strSrcPath & " -vf palettegen=max_colors=" & CStr(iColor) & ":reserve_transparent=1:stats_mode=full " & strPalettePath
  Rem paletteuse commandline
  cmd2= strFFPath & " -y -i " & strSrcPath & " -i " & strPalettePath & " -r " & CStr(iFPS) & " -lavfi paletteuse=dither=" & aDitherMethod(iMethod) & " -loop " & CStr(iLoop) & " " & strGifPath
  Rem gifsicle commandline
  cmd3= strGifsiclePath & " -V -U " & strTransparent & " -O3 " & strGifPath & " -o " & strGifOptPath
  
  If isCS Then
  objStdOut.WriteLine("Generating color palette using commandline:")
  objStdOut.WriteLine(cmd1)
  objStdOut.WriteBlankLines(2)
  Rem FFmpeg pass1
  Set oExec= objShell.Exec(cmd1)
  Do While oExec.Status = 0
    WScript.Sleep(100)
    objStdOut.Write(oExec.StdOut.ReadAll())
    WScript.StdErr.Write(oExec.StdErr.ReadAll())
  Loop
  Else
  Call objShell.Popup("Generating color palette..." & vbCrLf & cmd1,2)
  Call objShell.Run(cmd1, 1, True)
  End If
  
  If isCS Then
  objStdOut.WriteLine("Generating GIF using commandline:")
  objStdOut.WriteLine(cmd2)
  objStdOut.WriteBlankLines(2)
  Rem FFmpeg pass2
  Set oExec= objShell.Exec(cmd2)
  Do While oExec.Status = 0
    WScript.Sleep(100)
    objStdOut.Write(oExec.StdOut.ReadAll())
    WScript.StdErr.Write(oExec.StdErr.ReadAll())
  Loop
  Else
  Call objShell.Popup("Generating GIF..." & vbCrLf & cmd2,2)
  Call objShell.Run(cmd2, 1, true)
  End If
  
  If isCS Then
  objStdOut.WriteLine("Optimizing GIF using commandline:")
  objStdOut.WriteLine(cmd3)
  objStdOut.WriteBlankLines(2)
  Rem Gifsicle
  Set oExec= objShell.Exec(cmd3)
  Do While oExec.Status = 0
    WScript.Sleep(100)
    objStdOut.Write(oExec.StdOut.ReadAll())
    WScript.StdErr.Write(oExec.StdErr.ReadAll())
  Loop
  Else
  Call objShell.Popup("Optimizing GIF..." & vbCrLf & cmd3,2)
  Call objShell.Run(cmd3, 1, true)
  End If
  
  Rem delete palette image
  strShortPalette= Mid(strPalettePath, 2, Len(strPalettePath)-2)
  If (objFS.FileExists(strShortPalette)) Then
    objFS.DeleteFile(strShortPalette)
  End if  
  Next 
Else
  If isCS Then
  objStdOut.WriteLine("ffgif.vbs [file1] [file2][...fileN]" & vbCrLf & "Needs FFmpeg.exe and gifsicle.exe in the same folder as this script.")
  Else
  WScript.Echo("Please Drag-and-Drop video files on to the script's icon")
  End If
End If

If isCS Then    
objStdOut.WriteLine("Closing in 5 seconds.....")
WScript.Sleep(5000)
objStdOut.Close()
Else
Call objShell.Popup("Done! Closing in 5 seconds.....", 5)
End If
