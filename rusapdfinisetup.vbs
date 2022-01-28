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
' Russian prompts. File must be saved in ANSI encoding (windows-1251).
' UTF-8 will make prompts unreadable.
'
' Description END

' BEGINNING OF SCRIPT
Dim choiceClean, choiceLinearize, choiceEncrypt, choicePassLength, choiceSavePassToTxt, choiceAnnotate, choiceAes, objFso, strScriptDir, objIniFile
choiceClean = 6
choiceLinearize = 7
choiceEncrypt = 7
choicePassLength = 8
choiceSavePassToTxt = 7
choiceAnnotate = 6
choiceAes = 7
choiceClean = MsgBox("Очищать метаданные?", 4, "pdfClean")
choiceLinearize = MsgBox("Преобразовывать структуру документа в линейную?", 4, "pdfLinearize")
choiceEncrypt = MsgBox("Применять шифрование?", 4, "pdfEncrypt")
If choiceEncrypt = 6 Then
    choicePassLength = InputBox("Введите длину пароля" & vbCrLf & "(целое число между 1 и 255)", "pdfPassLength", "8")
    If IsNumeric(choicePassLength) Then
        If choicePassLength <= 0 Or choicePassLength > 255 Then
            choicePassLength = 8
        End If
        choicePassLength = CByte(choicePassLength)
    Else
        choicePassLength = 8
    End If
    choiceSavePassToTxt = MsgBox("Записывать пароль в файл?" & _
    vbCrLf & "(будет перезаписан при каждом запуске aadfPDF.vbs)", 4, "pdfSavePassToTxt")
    choiceAnnotate = MsgBox("Разрешить добавлять примечания?", 4, "pdfAnnotate")
    choiceAes = MsgBox("Шифровать методом AES?", 4, "pdfAes")
End If
Set objFso = CreateObject("Scripting.FileSystemObject")
strScriptDir = objFso.GetParentFolderName(WScript.ScriptFullName)
Set objIniFile = objFSO.CreateTextFile(strScriptDir & "\" & "apdfPDF.ini", True)
If choiceClean = 6 Then
    Call objIniFile.Write("pdfClean = 1" & vbCrLf)
Else
    Call objIniFile.Write("pdfClean = 0" & vbCrLf)
End If
If choiceLinearize = 6 Then
    Call objIniFile.Write("pdfLinearize = 1" & vbCrLf)
Else
    Call objIniFile.Write("pdfLinearize = 0" & vbCrLf)
End If
If choiceEncrypt = 6 Then
    Call objIniFile.Write("pdfEncrypt = 1" & vbCrLf)
Else
    Call objIniFile.Write("pdfEncrypt = 0" & vbCrLf)
End If
Call objIniFile.Write("pdfPassLength = " & choicePassLength & vbCrLf)
If choiceSavePassToTxt = 6 Then
    Call objIniFile.Write("pdfSavePassToTxt = 1" & vbCrLf)
Else
    Call objIniFile.Write("pdfSavePassToTxt = 0" & vbCrLf)
End If
If choiceAnnotate = 6 Then
    Call objIniFile.Write("pdfAnnotate = 1" & vbCrLf)
Else
    Call objIniFile.Write("pdfAnnotate = 0" & vbCrLf)
End If
If choiceAes = 6 Then
    Call objIniFile.Write("pdfAes = 1" & vbCrLf)
Else
    Call objIniFile.Write("pdfAes = 0" & vbCrLf)
End If
Call objIniFile.Write(vbCrLf)
Call objIniFile.Close()
Call MsgBox("Изменения в ini файл были успешно внесены!", 0, "apdfPDF.ini")
Call WScript.Quit()

' END OF SCRIPT

