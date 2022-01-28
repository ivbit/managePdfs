Option Explicit

' Intellectual property information START
' 
' Copyright (c) 2019 Ivan Bityutskiy 
' 
' Permission to use, copy, modify, and distribute this software for any
' purpose with or without fee is hereby granted, provided that the above
' copyright notice and this permission notice appear in all copies.
' 
' THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
' WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
' MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
' ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
' WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
' ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
' OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
' 
' Intellectual property information END

' Description START
'
' The script processes pdf files according to the settings from .ini file.
'
' Description END

' BEGINNING OF SCRIPT
Dim objFso, strScriptDir, objFolder, objFile, strPdfNames, arrAllPdfs, numIniPresent
Set objFso = CreateObject("Scripting.FileSystemObject")
strScriptDir = objFso.GetParentFolderName(WScript.ScriptFullName)
Set objFolder = objFSO.GetFolder(strScriptDir)
strPdfNames = ""
numIniPresent = 0
For Each objFile In objFolder.Files
    If objFile.Name = "apdfPDF.ini" Then
        numIniPresent = 1
    End If
    If (InStr(objFile.Name, ".") > 0) Then
        If (LCase(Mid(objFile.Name, InStrRev(objFile.Name, "."))) = ".pdf") Then
            strPdfNames = strPdfNames & objFile.Name & "/"
        End if 
    End If 
Next
If numIniPresent = 0 Then
    Call MsgBox("Config file apdfPDF.ini is not found!" & _
    vbCrLf & _
    "Launch apdfinisetup.vbs to generate new config file.", 0, "WScript.Quit")
    Call WScript.Quit()
End If
If strPdfNames = "" Then
    Call MsgBox("No PDF files found!", 0, "WScript.Quit")
    Call WScript.Quit()
End If
strPdfNames = Left(strPdfNames, Len(strPdfNames) - 1)
arrAllPdfs = Split(strPdfNames, "/")
Dim pdfClean, pdfLinearize, pdfEncrypt, pdfPassLength, pdfSavePassToTxt, pdfAnnotate, pdfAes
Dim objExtFile, arrSymbols, objWshShell, strComputerName, strProcessQpdf, strProcessExiftool
Dim strFileContent, strPassword, objPassFile, strThirdStep, strPdfFileName, strPdfFileQName, strArrItem
Dim numMax, numMin, numCounter, errUserPass
Set objExtFile = objFso.OpenTextFile(strScriptDir & "\" & "apdfPDF.ini", 1)
strFileContent = objExtFile.ReadAll()
Call objExtFile.Close()
Call ExecuteGlobal(strFileContent)
strPassword = ""
If pdfEncrypt = 1 Then
    arrSymbols = Array("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z","!","@","#","$","%","^","&","*","(", ")","-","+","0","1","2","3","4","5","6","7","8","9")
    numMax = 73
    numMin = 0
    For numCounter = 1 To pdfPassLength
        Call Randomize()
        strPassword = strPassword & arrSymbols(Int((numMax - numMin + 1) * Rnd + numMin))
    Next
    If pdfSavePassToTxt = 1 Then
        Set objPassFile = objFSO.CreateTextFile(strScriptDir & "\" & "aownerpass.txt", True)
        Call objPassFile.Write(strPassword & vbCrLf & vbCrLf)
        Call objPassFile.Close()
    End If
End If
strThirdStep = "qpdf.exe "
If pdfLinearize = 1 Then
    strThirdStep = strThirdStep & "--linearize "
End If
If pdfEncrypt = 1 Then
    strThirdStep = strThirdStep & "--encrypt """" """ & strPassword & """ 128 --print=full "
    If pdfAnnotate = 1 Then
        strThirdStep = strThirdStep & "--modify=annotate --extract=n "
    Else
        strThirdStep = strThirdStep & "--modify=none --extract=n "
    End If
    If pdfAes = 1 Then
        strThirdStep = strThirdStep & "--use-aes=y -- "
    Else
        strThirdStep = strThirdStep & "-- "
    End If
End If
Set objWshShell = CreateObject("WScript.Shell")
objWshShell.CurrentDirectory = strScriptDir
strProcessQpdf = "qpdf.exe"
strProcessExiftool = "exiftool.exe"
For Each strArrItem In arrAllPdfs
    strPdfFileName = strArrItem
    strPdfFileQName = "S" & strPdfFileName
    Call objWshShell.Run("qpdf.exe " & strPdfFileName & " " & strPdfFileQName, 0, True)
    errUserPass = 1
    For Each objFile In objFolder.Files
        If (InStr(objFile.Name, strPdfFileQName) > 0) Then
            errUserPass = 0
        End If
    Next
    If errUserPass = 1 Then
        Call MsgBox("The original file """ & _
        strPdfFileName & _
        """ cannot be processed! Unknown User Password!", 0, "WScript.Quit")
        Call WScript.Quit()
    End If
    Call objFso.DeleteFile(strPdfFileName)
    If pdfClean = 1 Then
        Call objWshShell.Run("exiftool.exe -all:all= " & strPdfFileQName, 0, True)
    End If
    Call objWshShell.Run(strThirdStep & strPdfFileQName & " " & strPdfFileName, 0, True)
    For Each objFile In objFolder.Files
        If (InStr(objFile.Name, "_original") > 0) Then
            Call objFso.DeleteFile(strPdfFileQName & "_original")
        End If
    Next
    Call objFso.DeleteFile(strPdfFileQName)
Next
MsgBox "Process completed successfully!", 0, "WScript.Quit"
Call WScript.Quit()

' END OF SCRIPT

