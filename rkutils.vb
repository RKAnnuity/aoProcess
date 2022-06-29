
Imports System
Imports System.Configuration
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlTypes
Imports System.Data.SqlDbType
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Web
Imports Microsoft.VisualBasic

#Const DB2 = 1
#If DB2 = 1 Then
Imports IBM.Data.DB2
#End If

Public Module rkutils

    Sub Wait(ByVal v_sec)
        Try
            Dim start As Double
            Dim finish As Double
            Dim totaltime As Double

            If TimeOfDay >= #11:59:55 PM# Then
                finish = 0
            End If
            start = Microsoft.VisualBasic.DateAndTime.Timer
            finish = start + v_sec   ' Set end time for v_sec seconds.
            Do While Microsoft.VisualBasic.DateAndTime.Timer < finish
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(200)
            Loop
            totaltime = Microsoft.VisualBasic.DateAndTime.Timer - start
        Catch ex As Exception
            MSG_ERROR("Wait", ex.ToString)
        End Try
    End Sub

    Public Sub DoEvents()
        Try
            System.Windows.Forms.Application.DoEvents()
        Catch ex As Exception
            MSG_ERROR("DoEvents", ex.ToString)
        End Try
    End Sub

    Public Function STR_TRIM(ByVal tSTRin As String, ByVal iLEN As Integer) As String
        Try
            If tSTRin.Length >= iLEN Then
                Return tSTRin.Substring(0, iLEN).Trim
            End If
            Return tSTRin
        Catch ex As Exception

        End Try
        Return ""
    End Function

    Public Function STR_LEFT(ByVal tSTRin As String, ByVal iLEN As Integer) As String
        Try
            If tSTRin.Length >= iLEN Then
                Return tSTRin.Substring(0, iLEN).Trim
            End If
            Return tSTRin
        Catch ex As Exception

        End Try
        Return ""
    End Function

    Public Function STR_RIGHT(ByVal tSTRin As String, ByVal iLEN As Integer) As String
        If tSTRin.Length >= iLEN Then
            Return tSTRin.Substring(tSTRin.Length - iLEN, iLEN).Trim
        End If
        Return tSTRin
    End Function

    Public Function STR_trim_CRLF(ByVal tLineIn As String) As String
        Dim iC As Integer
        Dim tLineOut As String = ""
        Try
            If tLineIn.Contains(vbCr) Or tLineIn.Contains(vbCrLf) Or tLineIn.Contains(Chr(10)) Or tLineIn.Contains(Chr(13)) Then
                For iC = 0 To tLineIn.Length - 1
                    Select Case tLineIn.Substring(iC, 1)
                        Case vbCr, vbCrLf, Chr(10), Chr(13)
                            tLineOut += " "
                        Case Else
                            tLineOut += tLineIn.Substring(iC, 1)
                    End Select
                Next
                Return tLineOut
            Else
                Return tLineIn
            End If
        Catch ex As Exception
            MSG_ERROR("STR_trim_CRLF", tLineOut + vbCr + vbCr + ex.ToString)
            Return ""
        End Try
    End Function

    Public Function STR_convert_AMP(ByVal tLineIn As String) As String
        Try
            Dim tLineOut As String = tLineIn
            tLineOut = tLineOut.Replace("&#39;", "'")
            tLineOut = tLineOut.Replace("&quot;", Chr(34))
            tLineOut = tLineOut.Replace("&nbsp;", " ")
            tLineOut = tLineOut.Replace("&lt;", "<")
            tLineOut = tLineOut.Replace("&gt;", ">")
            tLineOut = tLineOut.Replace("&amp;", "&")
            Return tLineOut
        Catch ex As Exception
            MSG_ERROR("STR_convert_AMP", ex.ToString)
            Return ""
        End Try
    End Function

    Public Function STR_NORMALIZE(ByVal tLineIn As String) As String
        ', ByVal tReplaceWith As String
        Try
            Dim tLineOut As String = tLineIn
            tLineOut = tLineOut.Replace("'", "")
            tLineOut = tLineOut.Replace(vbCr, "")
            tLineOut = tLineOut.Replace(vbLf, "")
            tLineOut = tLineOut.Replace(vbCrLf, "")
            Return tLineOut
        Catch ex As Exception
            MSG_ERROR("STR_NORMALIZE", ex.ToString)
            Return ""
        End Try
    End Function

    Public Function STR_SQLize(ByVal tLineIn As String) As String
        Try
            Dim tLineOut As String = tLineIn
            '*************************************
            '* 2013-01-31 RFK: changed to -
            tLineOut = tLineOut.Replace("'", "-")

            tLineOut = tLineOut.Replace("/", "-")
            tLineOut = tLineOut.Replace("\", "-")
            tLineOut = tLineOut.Replace("<", "[")
            tLineOut = tLineOut.Replace(">", "]")
            Return tLineOut
        Catch ex As Exception
            MSG_ERROR("STR_SQLize", ex.ToString)
            Return ""
        End Try
    End Function

    Public Function STR_TRAN(ByVal tLineIn As String, ByVal tFind As String, ByVal tReplace As String) As String
        Try
            Dim tLineOut As String = tLineIn
            tLineOut = tLineOut.Replace(tFind, tReplace)
            Return tLineOut
        Catch ex As Exception
            MSG_ERROR("STR_convert_AMP", ex.ToString)
            Return ""
        End Try
    End Function

    Public Function STR_LETTTERSONLY(ByVal tLineIn As String) As String
        Try
            Dim iC As Integer
            Dim tLineOut As String = ""
            Try
                For iC = 0 To tLineIn.Length - 1
                    Select Case tLineIn.Substring(iC, 1)
                        Case "A" To "Z"
                            tLineOut += tLineIn.Substring(iC, 1)
                        Case "a" To "z"
                            tLineOut += tLineIn.Substring(iC, 1)
                        Case " "
                            tLineOut += tLineIn.Substring(iC, 1)
                        Case Else
                            tLineOut += ""
                    End Select
                Next
                Return tLineOut
            Catch ex As Exception
                MSG_ERROR("STR_trim_CRLF", tLineOut + vbCr + vbCr + ex.ToString)
                Return ""
            End Try
        Catch ex As Exception

        End Try
    End Function

    Public Function STR_BREAK(ByVal tSTRin As String, ByVal FirstOr2nd As Integer) As String
        Try
            If tSTRin.Contains(" ") Then
                If FirstOr2nd = 1 Then
                    Return tSTRin.Substring(0, tSTRin.IndexOf(" ")).Trim
                Else
                    Return tSTRin.Substring(tSTRin.IndexOf(" ")).Trim
                End If
            End If
            Return tSTRin
        Catch ex As Exception

        End Try
    End Function

    Public Function STR_BREAK_AT(ByVal tSTRin As String, ByVal FirstOr2nd As Integer, ByVal tBreakCharacter As String) As String
        Try
            If tSTRin.Contains(tBreakCharacter) Then
                If FirstOr2nd = 1 Then
                    Return tSTRin.Substring(0, tSTRin.IndexOf(tBreakCharacter)).Trim
                Else
                    Return tSTRin.Substring(tSTRin.IndexOf(tBreakCharacter) + 1).Trim
                End If
            End If
            Return tSTRin
        Catch ex As Exception
            MSG_ERROR("STR_BREAK_AT", ex.ToString)
        End Try
        Return ""
    End Function

    Public Function STR_BREAK_PIECES(ByVal tLineIn As String, ByVal WhichOne As Integer, ByVal tBreakCharacter As String) As String
        Dim iC As Integer, iLast As Integer = -1, iBreakCTR As Integer = 1
        Dim inQUOTE As Boolean = False, swOK As Boolean = False
        Dim tLineOut As String = ""
        Try
            For iC = 0 To tLineIn.Length - 1
                swOK = False
                If tBreakCharacter = "," Then
                    If tLineIn.Substring(iC, 1) = Chr(34) Then
                        If inQUOTE = False Then
                            inQUOTE = True
                        Else
                            inQUOTE = False
                        End If
                    End If
                    If inQUOTE = False Then swOK = True
                Else
                    swOK = True
                End If
                '******************
                If swOK Then
                    If tLineIn.Substring(iC, 1) = tBreakCharacter Then
                        If iBreakCTR = WhichOne Then
                            Dim t2 As String = tLineIn.Substring(iLast + 1, iC - iLast - 1).Trim
                            If t2.Length = 0 Then t2 = "_" 'Chr(160) 'So Not Blank
                            Return t2
                        Else
                            iLast = iC
                        End If
                        iBreakCTR += 1
                    End If
                End If
            Next
            'MsgBox(tLineIn + vbCr + WhichOne.ToString + vbCr + iBreakCTR.ToString + vbCr + iLast.ToString + vbCr + iC.ToString + vbCr + iBreakCTR.ToString + vbCr)
            If iBreakCTR = WhichOne Then Return tLineIn.Substring(iLast + 1).Trim 'The Last One
            Return ""
        Catch ex As Exception
            MSG_ERROR("STR_BREAK_PIECES", tLineOut + vbCr + vbCr + ex.ToString)
            Return ""
        End Try
    End Function

    Public Function STR_BREAK_STR(ByVal tSTRin As String, ByVal tStartString As String, ByVal tStopString As String, ByVal iAfterStart As Integer) As String
        Try
            If tSTRin.Contains(tStartString) And tSTRin.Contains(tStopString) Then
                Dim i1 As Integer = tSTRin.IndexOf(tStartString)
                Dim i2 As Integer = tSTRin.IndexOf(tStopString)
                If iAfterStart > 0 Then i1 = i1 + tStartString.Length
                If i1 >= 0 And i2 > 0 And i2 - i1 <= tSTRin.Length Then Return tSTRin.Substring(i1, i2 - i1).Trim
            End If
            Return tSTRin
        Catch ex As Exception
            MSG_ERROR("STR_BREAK_AT", ex.ToString)
        End Try
        Return ""
    End Function

    Public Function SecondsToTime(ByVal sttSeconds As Long, ByVal Num2Return As Integer) As String
        Try
            Dim tHour As String, tMin As String, tSec As String
            Dim dN As Double, dM As Double, dS As Double
            Dim TmpH As Integer, tmpM As Integer, TmpS As Integer

            dN = sttSeconds
            'if(Average) if(OCtr > 0) TmpN=(TmpN/OCtr);
            TmpH = 0
            dM = (dN / 60)
            dS = dN Mod 60

            Select Case dM
                Case 60 To 117
                    dM = dM - 60
                    TmpH = 1
                Case 119 To 176
                    dM = dM - 120
                    TmpH = 2
                Case 178 To 235
                    dM = dM - 180
                    TmpH = 3
                Case 237 To 294
                    dM = dM - 240
                    TmpH = 4
                Case 296 To 353
                    dM = dM - 300
                    TmpH = 5
                Case 355 To 412
                    dM = dM - 360
                    TmpH = 6
                Case 414 To 471
                    dM = dM - 420
                    TmpH = 7
                Case 473 To 530
                    dM = dM - 480
                    TmpH = 8
                Case 532 To 589
                    dM = dM - 540
                    TmpH = 9
                Case 591 To 648
                    dM = dM - 600
                    TmpH = 10
                Case 650 To 707
                    dM = dM - 660
                    TmpH = 11
                Case 709 To 766
                    dM = dM - 720
                    TmpH = 12
                Case 768 To 825
                    dM = dM - 780
                    TmpH = 13
                Case 827 To 882
                    dM = dM - 840
                    TmpH = 14
                Case 884 To 941
                    dM = dM - 900
                    TmpH = 15
                Case 943 To 1000
                    dM = dM - 960
                    TmpH = 16
                Case 1002 To 1059
                    dM = dM - 1020
                    TmpH = 17
                Case 1061 To 1118
                    dM = dM - 1080
                    TmpH = 18
                Case 1120 To 1177
                    dM = dM - 1140
                    TmpH = 19
                Case 1179 To 1236
                    dM = dM - 1200
                    TmpH = 20
                Case 1238 To 1295
                    dM = dM - 1260
                    TmpH = 21
                Case 1297 To 1354
                    dM = dM - 1320
                    TmpH = 22
                Case 1356 To 1413
                    dM = dM - 1380
                    TmpH = 23
                Case 1415 To 1472
                    dM = dM - 1440
                    TmpH = 24
                Case 1474 To 9999
                    TmpH = -1
            End Select

            tmpM = Fix(Str(dM))
            TmpS = Fix(Str(dS))
            If ((TmpS < 1) Or (TmpS > 59)) Then TmpS = 0
            If ((tmpM < 1) Or (tmpM > 59)) Then tmpM = 0
            If TmpH < 0 Then
                tHour = "++"
            Else
                If TmpH < 10 Then
                    tHour = "0" + Trim(Str(TmpH))
                Else
                    tHour = Trim(Str(TmpH))
                End If
            End If
            If tmpM < 10 Then
                tMin = "0" + Trim(Str(tmpM))
            Else
                tMin = Trim(Str(tmpM))
            End If
            If TmpS < 10 Then
                tSec = "0" + Trim(Str(TmpS))
            Else
                tSec = Trim(Str(TmpS))
            End If
            Select Case Num2Return
                Case 5
                    SecondsToTime = tMin + ":" + tSec
                Case 8
                    SecondsToTime = tHour + ":" + tMin + ":" + tSec
                Case Else
                    SecondsToTime = tHour + ":" + tMin + ":" + tSec
            End Select
        Catch ex As Exception
        End Try
    End Function

    Public Function TimeToSeconds(ByVal TIMEin As String) As Long
        Try
            Dim dH As Double, dM As Double, dS As Double
            Dim TmpH As Integer, tmpM As Integer, TmpS As Integer

            If Len(Trim(TIMEin)) = 8 And Mid(TIMEin, 3, 1) = ":" Then
                TmpH = Mid(TIMEin, 1, 2)
                tmpM = Mid(TIMEin, 4, 2)
                TmpS = Mid(TIMEin, 7, 2)

                dH = (Val(TmpH) * 60) * 60
                dM = Val(tmpM) * 60
                dS = Val(TmpS)
                TimeToSeconds = dH + dM + dS
                Exit Function
            End If
        Catch ex As Exception

        End Try
        Return 0
    End Function

    Public Function TimeSecondsElapsed(ByVal timeInElapsed As String) As Long
        Try
            Dim dStart As Double, dNow As Double

            dStart = TimeToSeconds(timeInElapsed)
            'dNow = TimeToSeconds(Time24(8))

            TimeSecondsElapsed = dNow - dStart
        Catch ex As Exception

        End Try
        Return 0
    End Function

    Public Function DateDaysInMonth(ByVal dmMonth As Integer) As Integer
        Try
            Select Case dmMonth
                Case 1  'Jan
                    DateDaysInMonth = 31
                Case 2  'Feb
                    DateDaysInMonth = 28
                Case 3  'Mar
                    DateDaysInMonth = 31
                Case 4  'Apr
                    DateDaysInMonth = 30
                Case 5  'May
                    DateDaysInMonth = 31
                Case 6  'June
                    DateDaysInMonth = 30
                Case 7  'July
                    DateDaysInMonth = 31
                Case 8  'Aug
                    DateDaysInMonth = 31
                Case 9  'Sep
                    DateDaysInMonth = 30
                Case 10 'Oct
                    DateDaysInMonth = 31
                Case 11 'Nov
                    DateDaysInMonth = 30
                Case 12 'Dec
                    DateDaysInMonth = 31
                Case Else
                    DateDaysInMonth = 0
            End Select
        Catch ex As Exception

        End Try
        Return 0
    End Function

    Public Function DateToday(ByVal Num2Return) As String
        Try

            '2008-05-20 RFK: This is a copy from rklib.vb
            Dim tSTR As String
            tSTR = ""
            Select Case Num2Return
                Case 8
                    tSTR = Now.Year.ToString
                    If Now.Month >= 10 Then
                        tSTR = tSTR + Now.Month.ToString
                    Else
                        tSTR = tSTR + "0" + Now.Month.ToString
                    End If
                    If Now.Day >= 10 Then
                        tSTR = tSTR + Now.Day.ToString
                    Else
                        tSTR = tSTR + "0" + Now.Day.ToString
                    End If
                Case 16     'ccyymmddHHMMSSss
                    tSTR = Now.Year.ToString
                    If Now.Month >= 10 Then
                        tSTR = tSTR + Now.Month.ToString
                    Else
                        tSTR = tSTR + "0" + Now.Month.ToString
                    End If
                    If Now.Day >= 10 Then
                        tSTR = tSTR + Now.Day.ToString
                    Else
                        tSTR = tSTR + "0" + Now.Day.ToString
                    End If
                    If Now.Hour >= 10 Then
                        tSTR = tSTR + Now.Hour.ToString
                    Else
                        tSTR = tSTR + "0" + Now.Hour.ToString
                    End If
                    If Now.Minute >= 10 Then
                        tSTR = tSTR + Now.Minute.ToString
                    Else
                        tSTR = tSTR + "0" + Now.Minute.ToString
                    End If
                    If Now.Second >= 10 Then
                        tSTR = tSTR + Now.Second.ToString
                    Else
                        tSTR = tSTR + "0" + Now.Second.ToString
                    End If
                    tSTR = tSTR + Now.Millisecond.ToString
                Case 18 'ccyy-MM-dd HH:mm:ss
                    tSTR = Now.Year.ToString
                    tSTR = tSTR + "-"
                    If Now.Month >= 10 Then
                        tSTR = tSTR + Now.Month.ToString
                    Else
                        tSTR = tSTR + "0" + Now.Month.ToString
                    End If
                    tSTR = tSTR + "-"
                    If Now.Day >= 10 Then
                        tSTR = tSTR + Now.Day.ToString
                    Else
                        tSTR = tSTR + "0" + Now.Day.ToString
                    End If
                    tSTR = tSTR + " "
                    If Now.Hour >= 10 Then
                        tSTR = tSTR + Now.Hour.ToString
                    Else
                        tSTR = tSTR + "0" + Now.Hour.ToString
                    End If
                    tSTR = tSTR + ":"
                    If Now.Minute >= 10 Then
                        tSTR = tSTR + Now.Minute.ToString
                    Else
                        tSTR = tSTR + "0" + Now.Minute.ToString
                    End If
                    tSTR = tSTR + ":"
                    If Now.Second >= 10 Then
                        tSTR = tSTR + Now.Second.ToString
                    Else
                        tSTR = tSTR + "0" + Now.Second.ToString
                    End If
                Case 20 'mm/dd/ccyy HH:MM:SS
                    tSTR = ""
                    If Now.Month >= 10 Then
                        tSTR += Now.Month.ToString
                    Else
                        tSTR += "0" + Now.Month.ToString
                    End If
                    tSTR += "-"
                    If Now.Day >= 10 Then
                        tSTR += Now.Day.ToString
                    Else
                        tSTR += "0" + Now.Day.ToString
                    End If
                    tSTR += "-"
                    tSTR += Now.Year.ToString
                    tSTR += " "
                    If Now.Hour >= 10 Then
                        tSTR += Now.Hour.ToString
                    Else
                        tSTR += "0" + Now.Hour.ToString
                    End If
                    tSTR += ":"
                    If Now.Minute >= 10 Then
                        tSTR += Now.Minute.ToString
                    Else
                        tSTR += "0" + Now.Minute.ToString
                    End If
                    tSTR += ":"
                    If Now.Second >= 10 Then
                        tSTR += Now.Second.ToString
                    Else
                        tSTR += "0" + Now.Second.ToString
                    End If
            End Select
            Return tSTR
        Catch ex As Exception

        End Try
    End Function

    Public Function TimeToday(ByVal Num2Return) As String
        Try
            '2009-01-26 RFK: This is a copy from rklib.vb
            Dim tSTR As String
            tSTR = ""
            Select Case Num2Return
                Case 5     'HH:MM   'SSss
                    If Now.Hour >= 10 Then
                        tSTR += Now.Hour.ToString
                    Else
                        tSTR += "0" + Now.Hour.ToString
                    End If
                    tSTR += ":"
                    If Now.Minute >= 10 Then
                        tSTR += Now.Minute.ToString
                    Else
                        tSTR += "0" + Now.Minute.ToString
                    End If
                Case 6      'HHmmss
                    If Now.Second >= 10 Then
                        tSTR = tSTR + Now.Second.ToString
                    Else
                        tSTR = tSTR + "0" + Now.Second.ToString
                    End If
                    tSTR = tSTR + Now.Millisecond.ToString
            End Select
            Return tSTR
        Catch ex As Exception

        End Try
    End Function

    Public Function FILE_delete(ByVal FileFullPath As String) As Boolean
        Try
            If My.Computer.FileSystem.FileExists(FileFullPath) Then
                My.Computer.FileSystem.DeleteFile(FileFullPath)
                If My.Computer.FileSystem.FileExists(FileFullPath) Then
                    Return False
                Else
                    Return True
                End If
            End If
        Catch ex As Exception

        End Try
        Return False
    End Function

    Public Function FILE_read(ByVal FileFullPath As String) As String
        Try
            If My.Computer.FileSystem.FileExists(FileFullPath) Then
                Return My.Computer.FileSystem.ReadAllText(FileFullPath)
            End If
        Catch ex As Exception

        End Try
        Return ""
    End Function

    Public Function FILE_contains(ByVal FileFullPath As String, ByVal tContains As String) As String
        Try
            Dim tSTR As String, nSTR As String
            If My.Computer.FileSystem.FileExists(FileFullPath) Then
                tSTR = My.Computer.FileSystem.ReadAllText(FileFullPath)
                If tSTR.Contains(tContains) Then
                    nSTR = tSTR.Substring(tSTR.IndexOf(tContains) + Len(tContains))
                    If nSTR.Contains(vbCr) Then
                        nSTR = nSTR.Substring(0, nSTR.IndexOf(vbCr))
                    End If
                    Return (nSTR)
                Else
                    Return ""
                End If
            End If
            Return ""
        Catch ex As Exception

        End Try
    End Function

    Function FILE_create(ByVal FileFullPath As String, ByVal OverWrite As Boolean, ByVal AppendToFile As Boolean, ByVal toWrite As String) As String
        Try
            If My.Computer.FileSystem.FileExists(FileFullPath) Then
                If OverWrite Then
                    FILE_delete(FileFullPath)
                    If My.Computer.FileSystem.FileExists(FileFullPath) Then
                        Return False
                    End If
                Else
                    If AppendToFile Then
                        My.Computer.FileSystem.WriteAllText(FileFullPath, toWrite, True)
                        Return True
                    Else
                        Return False
                    End If
                End If
            End If
            My.Computer.FileSystem.WriteAllText(FileFullPath, toWrite, False)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function LineFind(ByVal tSTR As String, ByVal tFind As String) As Integer
        Try
            Dim tLine As String
            Dim i1 As Integer, iLine As Integer

            tLine = ""
            iLine = 1
            i1 = 1
            Do While i1 < Len(tSTR)
                If Mid(tSTR, i1, 1) = Chr(10) Or i1 = Len(tSTR) + 1 Then
                    If tLine.Contains(tFind) Then
                        Return iLine
                    End If
                    If Mid(tSTR, i1 + 1, 1) = vbCr Then
                        i1 = i1 + 1
                    End If

                    tLine = ""
                    iLine = iLine + 1
                Else
                    If Mid(tSTR, i1, 1) = Chr(9) Then
                        tLine = tLine + " "
                    Else
                        tLine = tLine + Mid(tSTR, i1, 1)
                    End If
                End If
                i1 = i1 + 1
            Loop
            Return 0
        Catch ex As Exception

        End Try
    End Function

    Public Function LineRead(ByVal tSTR As String, ByVal iLineRead As Integer) As String
        Try
            Dim tLine As String
            Dim i1 As Integer, iLine As Integer

            tLine = ""
            iLine = 0
            i1 = 1
            Do While i1 < Len(tSTR)
                If Mid(tSTR, i1, 1) = Chr(10) Or i1 = Len(tSTR) + 1 Then
                    If iLine = iLineRead Then
                        Return tLine
                    End If
                    If Mid(tSTR, i1 + 1, 1) = vbCr Then
                        i1 = i1 + 1
                    End If
                    tLine = ""
                    iLine = iLine + 1
                Else
                    tLine = tLine + Mid(tSTR, i1, 1)
                End If
                i1 = i1 + 1
            Loop
            Return ""
        Catch ex As Exception

        End Try
    End Function

    Public Function STR_FromLine(ByVal tSTR As String, ByVal iLineRead As Integer) As String
        Try
            Dim tLine As String
            Dim i1 As Integer, i2 As Integer, iLine As Integer

            tLine = ""
            iLine = 0
            i1 = 1
            Do While i1 < Len(tSTR)
                If Mid(tSTR, i1, 1) = Chr(10) Or i1 = Len(tSTR) + 1 Then
                    If iLine = iLineRead Then
                        Return Mid(tSTR, i2)
                    End If
                    If Mid(tSTR, i1 + 1, 1) = vbCr Then
                        i1 = i1 + 1
                    End If
                    tLine = ""
                    i2 = i1
                    iLine = iLine + 1
                Else
                    tLine = tLine + Mid(tSTR, i1, 1)
                End If
                i1 = i1 + 1
            Loop
            Return ""
        Catch ex As Exception

        End Try
    End Function

    Public Function BreakSPAN(ByVal tTDline As String) As String
        Try
            Dim i1 As Integer = 0, i2 As Integer = 0, i3 As Integer = 0

            BreakSPAN = ""
            If tTDline.Contains("</SPAN") Then
                For i1 = tTDline.Length - 8 To 0 Step -1
                    If tTDline.Substring(i1, 7) = "</SPAN>" Then
                        i3 = i1 - 1
                    End If
                    If i3 > 0 Then
                        If tTDline.Substring(i1, 1) = ">" Then
                            i2 = i1
                            Exit For
                        End If
                    End If
                Next
            End If
            If i2 > 0 And i3 > 0 And i3 < tTDline.Length - 1 Then
                BreakSPAN = tTDline.Substring(i2 + 1, i3 - i2)
            End If
        Catch ex As Exception

        End Try
    End Function

    Public Function BreakTD(ByVal tTDline As String) As String
        Try
            Dim i1 As Integer, i2 As Integer
            Dim tNewLine As String

            i1 = InStr(1, tTDline, ">")
            If i1 > 0 Then
                i2 = InStr(i1, tTDline, "<")
                If i2 > 0 Then
                    tTDline = Mid(tTDline, i1 + 1, i2 - i1 - 1)
                Else
                    tTDline = Mid(tTDline, i1 + 1)
                End If
            Else
                tTDline = Mid(tTDline, i1 + 1)
            End If
            If tTDline = "&nbsp;" Then
                tTDline = ""
            End If
            If tTDline.Length > 0 Then
                If tTDline.Substring(0, 1) = "<" Or tTDline.Substring(0, 1) = ">" Then
                    tTDline = ""
                End If
                If tTDline.Contains("<") Then
                    'MsgBox(tTDline)
                End If
            End If
            tNewLine = ""
            For i1 = 0 To tTDline.Length - 1
                If tTDline.Length >= i1 + 6 Then
                    If tTDline.Substring(i1, 6) = "&nbsp;" Then
                        tNewLine = tNewLine + " "
                        i1 = i1 + 5
                    Else
                        If tTDline.Substring(i1, 1) >= " " And tTDline.Substring(i1, 1) <= "z" Then
                            tNewLine = tNewLine + tTDline.Substring(i1, 1)
                        End If
                    End If
                Else
                    If tTDline.Substring(i1, 1) >= " " And tTDline.Substring(i1, 1) <= "z" Then
                        tNewLine = tNewLine + tTDline.Substring(i1, 1)
                    Else
                        'MsgBox(tTDline.Substring(i1, 1))
                    End If
                End If
            Next
            Return tNewLine
        Catch ex As Exception
        End Try
    End Function

    Public Function ComboBox_SetValue(ByVal cCombo As ComboBox, ByVal tVal As String) As Integer
        Try
            Dim i2 As Integer
            For i2 = 0 To cCombo.Items.Count - 1
                If UCase(cCombo.Items(i2).ToString) = UCase(tVal) Then
                    cCombo.SelectedIndex = i2
                    Return i2
                End If
            Next
        Catch ex As Exception
            'MSG_warning(ex.ToString)
        End Try
        Return -1
    End Function

    Public Function DataGridViewContains(ByVal gGrid As DataGridView, ByVal tField As String, ByVal tValue As String) As Integer
        Try
            Dim i1Grid As Integer
            For i1Grid = 0 To gGrid.Rows.Count - 1
                If rkutils.DataGridView_ValueByColumnName(gGrid, tField, i1Grid).Trim = tValue Then Return i1Grid
            Next
        Catch ex As Exception

        End Try
        Return -1
    End Function

    Public Function DataGridView_ColumnByName(ByVal gGrid As DataGridView, ByVal tColName As String) As Integer
        Try
            Dim i2 As Integer
            For i2 = 0 To gGrid.ColumnCount - 1
                If UCase(gGrid.Columns(i2).Name) = UCase(tColName) Then
                    Return i2
                End If
            Next
        Catch ex As Exception
            'MSG_warning(ex.ToString)
        End Try
        Return -1
    End Function

    Public Function DataGridView_ValueByColumnName(ByVal gGrid As DataGridView, ByVal tColName As String, ByVal iRow As Integer) As String
        Try
            Dim rc1 As Integer = DataGridView_ColumnByName(gGrid, tColName)
            If rc1 >= 0 Then
                If Len(gGrid.Item(rc1, iRow).Value) > 0 Then
                    Return gGrid.Item(rc1, iRow).Value.ToString()
                Else
                    Return ""
                End If
            End If
        Catch ex As Exception
            'MSG_warning(ex.ToString)
        End Try
        Return ""
    End Function

    Public Function DataGridView_SetValueByColumnName(ByVal gGrid As DataGridView, ByVal tColName As String, ByVal iRow As Integer, ByVal tValue As String)
        Try
            Dim rc1 As Integer = DataGridView_ColumnByName(gGrid, tColName)
            If rc1 >= 0 Then
                gGrid.Item(rc1, iRow).Value = tValue
            End If
        Catch ex As Exception
            'MSG_warning(ex.ToString)
        End Try
        Return ""
    End Function

    Public Function DataGridView_AddColumn(ByVal gGrid As DataGridView, ByVal tColumnName As String)
        Try
            If DataGridView_ColumnByName(gGrid, tColumnName) >= 0 Then
                'Nothing
            Else
                gGrid.Columns.Add(tColumnName, tColumnName)
            End If
            Return True
        Catch ex As Exception
            'MSG_warning(ex.ToString)
        End Try
        Return False
    End Function

    Public Function WhoAmI() As String
        Try
            'Dim tUSR As String = HttpContext.Current.User.Identity.Name.ToString
            Dim tUSR As String = My.User.Name
            Dim lSlash As Integer = tUSR.LastIndexOf("\")
            If lSlash > 0 Then
                tUSR = tUSR.Substring(lSlash + 1)
            End If
            Return tUSR
        Catch ex As Exception

        End Try
    End Function

    Public Sub DB_COMMAND(ByVal tDB As String, ByVal SQLConnectionString As String, ByVal SQLuser As String, ByVal tCOMMAND As String)
        Try
            '******************************************************************
            '* 
            'eProcess.MsgStatus("DB_COMMAND:" + tCOMMAND, True)
            '******************************************************************
            Select Case tDB
                Case "DB2"
#If DB2 = 1 Then
                    Dim dbConnection As IBM.Data.DB2.iSeries.iDB2Connection = New IBM.Data.DB2.iSeries.iDB2Connection(SQLConnectionString + SQLuser)
                    Dim dbCommand As IBM.Data.DB2.iSeries.iDB2Command = New IBM.Data.DB2.iSeries.iDB2Command()
                    dbCommand.CommandText = tCOMMAND
                    dbCommand.Connection = dbConnection
                    dbCommand.CommandTimeout = 0
                    dbCommand.Connection.Open()
                    dbCommand.ExecuteNonQuery()

                    dbCommand.Dispose()
                    dbCommand = Nothing
                    dbConnection.Close()
                    dbConnection.Dispose()
                    dbConnection = Nothing
#End If
                Case "MSSQL"
                    Dim dbConnection As New SqlConnection(SQLConnectionString + SQLuser)        ' The SqlConnection class allows you to communicate with SQL Server.
                    Dim dbCommand As New SqlCommand(tCOMMAND, dbConnection)                ' A SqlCommand object is used to execute the SQL commands.
                    dbCommand.CommandText = tCOMMAND
                    dbCommand.Connection = dbConnection
                    dbCommand.CommandTimeout = 0
                    dbCommand.Connection.Open()
                    dbCommand.ExecuteNonQuery()
                    dbCommand.Dispose()
                    dbCommand = Nothing
                    dbConnection.Close()
                    dbConnection.Dispose()
                    dbConnection = Nothing
                Case "MYSQL"
                    'Dim dbConnection As New OdbcConnection(SQLConnectionString + SQLuser)    'The SqlConnection class allows you to communicate with SQL Server.
                    'Dim dbCommand As New OdbcCommand(tCOMMAND, dbConnection)            'A SqlCommand object is used to execute the SQL commands.
                    'dbCommand.CommandText = tCOMMAND
                    'dbCommand.Connection = dbConnection
                    'dbCommand.Connection.Open()
                    'dbCommand.ExecuteNonQuery()
                    'dbCommand.Dispose()
                    'dbCommand = Nothing
                    'dbConnection.Close()
                    'dbConnection.Dispose()
                    'dbConnection = Nothing
            End Select
        Catch ex As Exception
            MSG_ERROR("DB_COMMAND/" + tCOMMAND, ex.ToString)
        End Try
    End Sub

    Public Function COMMAND_STATUS(ByVal SQLConnectionString As String, ByVal SQLuser As String, ByVal tLOCX As String, ByVal tSTATUS As String, ByVal tCC As String, ByVal tRAC As String, ByVal tMC As String) As Boolean
        Try
            Dim SQLstring As String = "INSERT INTO RevMD.dbo.commands "
            SQLstring += " (COMMAND, LOCX, STATUS, CC, RAC, MC, MODIFIED_BY, MODIFIED_DATE, INSERT_DATE)"
            SQLstring += " VALUES("
            SQLstring += "'STATUS'"
            SQLstring += ", '" + tLOCX + "'"
            SQLstring += ", '" + tSTATUS + "'"
            SQLstring += ", '" + tCC + "'"
            SQLstring += ", '" + tRAC + "'"
            SQLstring += ", '" + tMC + "'"
            SQLstring += ", '" + WhoAmI() + "'"
            SQLstring += ", '" + Date.Now.ToString + "'"
            SQLstring += ", '" + Date.Now.ToString + "'"
            SQLstring += ")"
            DB_COMMAND("MSSQL", SQLConnectionString, SQLuser, SQLstring)
            Return True
        Catch ex As Exception
            'MSG_post("AD_memberofGroup:" + ex.ToString, "YELLOW", "RED")
        End Try
        Return False
    End Function

    Public Function SQL_READ_FIELD(ByVal gGrid As DataGridView, ByVal tDB As String, ByVal tFIELDNAME As String, ByVal tConnectionString As String, ByVal tSQLuser As String, ByVal tSELECTstring As String) As String
        Try
            If SQL_READ_DATAGRID(gGrid, tDB, tFIELDNAME, tConnectionString, tSQLuser, tSELECTstring) Then
                Return DataGridView_ValueByColumnName(gGrid, tFIELDNAME, 0)
            End If
            Return ""
        Catch ex As Exception
            MSG_ERROR("SQL_READ_FIELD", ex.ToString)
        End Try
        Return ""
    End Function

    Public Function SQL_READ_DATATABLE(ByVal dT As DataTable, ByVal tDB As String, ByVal tMODULE As String, ByVal tSQLConnectionString As String, ByVal tSQLuser As String, ByVal tSELECTstring As String) As Boolean
        Try
            Select Case tDB.ToUpper
                Case "DB2"
#If DB2 = 1 Then
                    Dim dbConnection As IBM.Data.DB2.iSeries.iDB2Connection = New IBM.Data.DB2.iSeries.iDB2Connection(tSQLConnectionString + tSQLuser)
                    Dim dbCommand As IBM.Data.DB2.iSeries.iDB2Command = New IBM.Data.DB2.iSeries.iDB2Command()
                    dbCommand.CommandText = tSELECTstring
                    dbCommand.Connection = dbConnection
                    dbCommand.CommandTimeout = 0

                    Dim dAdapter As IBM.Data.DB2.iSeries.iDB2DataAdapter = New IBM.Data.DB2.iSeries.iDB2DataAdapter
                    dAdapter.SelectCommand = dbCommand

                    Dim dataSet As System.Data.DataSet = New System.Data.DataSet
                    dAdapter.Fill(dT)

                    dbCommand.Dispose()
                    dbCommand = Nothing
                    dbConnection.Close()
                    dbConnection.Dispose()
                    dbConnection = Nothing
                    If dT.Rows.Count > 0 Then Return True
#End If
                Case "MSSQL"
                    Dim dbConnection As New SqlConnection(tSQLConnectionString + tSQLuser)        ' The SqlConnection class allows you to communicate with SQL Server.
                    Dim dbCommand As New SqlCommand(tSELECTstring, dbConnection)                ' A SqlCommand object is used to execute the SQL commands.

                    Dim dAdapter As New SqlDataAdapter(dbCommand)
                    Dim mDataSet As New DataSet()
                    dAdapter.Fill(dT)

                    dbCommand.Dispose()
                    dbCommand = Nothing
                    dbConnection.Close()
                    dbConnection.Dispose()
                    dbConnection = Nothing
                    If dT.Rows.Count > 0 Then Return True
                Case "MYSQL"
                    '
                Case "FOXPRO"
                    Dim dbConnection As New OleDbConnection("Provider=vfpoledb.1;Data Source=" + tSQLConnectionString + ";Collating Sequence=machine")
                    Dim dbCommand As New OleDbCommand
                    Dim dbDataAdapter As New OleDbDataAdapter

                    dbCommand.CommandText = tSELECTstring
                    dbCommand.Connection = dbConnection

                    dbDataAdapter.SelectCommand = dbCommand
                    dbDataAdapter.Fill(dT)

                    dbCommand.Dispose()
                    dbCommand = Nothing
                    dbConnection.Close()
                    dbConnection.Dispose()
                    dbConnection = Nothing
                    If dT.Rows.Count > 0 Then Return True
            End Select
        Catch ex As Exception
            Console.Write("SQL_READ_DATATABLE" + ex.ToString)
        End Try
        Return False
    End Function

    Public Function IS_File(ByVal FileFullPath As String) As Boolean
        Try
            If FileFullPath.Contains("*") Then
                'Dim tSTR As String = My.Computer.FileSystem.GetFiles(FileFullPath, FileIO.SearchOption.SearchTopLevelOnly, "*.*")
                'Dim tSTRING As string[] files = Directory.GetFiles("D:\Documents and Settings\Lou\My Documents\Visual Studio", "*.txt");
            End If
            Return My.Computer.FileSystem.FileExists(FileFullPath)
        Catch ex As Exception

        End Try
    End Function

    Public Function Listbox_Contains(ByVal lList As ListBox, ByVal tValue As String, ByVal bAdd As Boolean) As Boolean
        Try
            Dim tLIST As String = ""
            For i1ctr = 0 To lList.Items.Count - 1
                tLIST = lList.Items(i1ctr).ToString
                If tLIST = tValue Then
                    Return True
                Else
                    If tLIST.Contains(" ") Then
                        If tLIST.Substring(0, tLIST.IndexOf(" ")) = tValue Then
                            Return True
                        End If
                    End If
                End If
            Next
            If bAdd And tValue.Trim.Length > 0 Then
                lList.Items.Add(tValue)
            End If
            Return False
        Catch ex As Exception

        End Try
    End Function

    Public Function Listbox_Row(ByVal lList As ListBox, ByVal tValue As String, ByVal bAdd As Boolean) As Integer
        Try
            Dim tLIST As String = ""
            For i1ctr = 0 To lList.Items.Count - 1
                tLIST = lList.Items(i1ctr).ToString
                If tLIST = tValue Then
                    Return i1ctr
                Else
                    If tLIST.Contains(tValue) Then
                        Return i1ctr
                    End If
                End If
            Next
            If bAdd And tValue.Trim.Length > 0 Then
                lList.Items.Add(tValue)
                Return lList.Items.Count - 1
            End If
            Return -1
        Catch ex As Exception

        End Try
    End Function

    Public Function Listbox_Text_Contains(ByVal lList As ListBox, ByVal tValue As String, ByVal bAdd As Boolean) As Boolean
        '*************************************************************************************
        '* 2012-01-10 RFK: If any of the WORDS within the tValue are in lList then return TRUE
        Try
            Dim i1ctr As Integer, i2ctr As Integer = 0
            If tValue.Contains(" ") Then
                For i1ctr = 0 To tValue.Length
                    If tValue.Substring(i1ctr, 1) = " " Then
                        If Listbox_Contains(lList, tValue.Substring(i2ctr, i1ctr - i2ctr), False) Then Return True
                        If Listbox_Contains(lList, UCase(tValue.Substring(i2ctr, i1ctr - i2ctr)), False) Then Return True
                        i2ctr = i1ctr + 1
                    End If
                Next
                'Check it 1 more time for end of line
                If Listbox_Contains(lList, tValue.Substring(i2ctr, i1ctr - i2ctr), False) Then Return True
                If Listbox_Contains(lList, UCase(tValue.Substring(i2ctr, i1ctr - i2ctr)), False) Then Return True
            Else
                Return Listbox_Contains(lList, tValue, False)
            End If
            If bAdd And tValue.Trim.Length > 0 Then
                lList.Items.Add(tValue)
            End If
        Catch ex As Exception
            '
        End Try
        Return False
    End Function

    Public Function Listbox_Value(ByVal lList As ListBox, ByVal tValue As String, ByVal BreakSpace As Boolean) As String
        Try
            Dim tLIST As String = ""
            For i1ctr = 0 To lList.Items.Count - 1
                tLIST = lList.Items(i1ctr).ToString
                If tLIST = tValue Then
                    Return tLIST
                Else
                    If BreakSpace Then
                        If tLIST.Contains(" ") Then
                            If tLIST.Substring(0, tLIST.IndexOf(" ")) = tValue Then
                                Return tLIST
                            End If
                        End If
                    End If
                End If
            Next
            Return ""
        Catch ex As Exception

        End Try
    End Function

    Public Function Listbox_Select(ByVal lList As ListBox, ByVal tValue As String, ByVal BreakSpace As Boolean) As Integer
        Try
            Dim tLIST As String = ""
            For i1ctr = 0 To lList.Items.Count - 1
                tLIST = lList.Items(i1ctr).ToString
                If tLIST = tValue Then
                    lList.SelectedIndex = i1ctr
                    Return i1ctr
                Else
                    If BreakSpace Then
                        If tLIST.Contains(" ") Then
                            If tLIST.Substring(0, tLIST.IndexOf(" ")) = tValue Then
                                lList.SelectedIndex = i1ctr
                                Return i1ctr
                            End If
                        End If
                    End If
                End If
            Next
            Return -1
        Catch ex As Exception

        End Try
    End Function

    Public Function NOTES_MAXNUMBER(ByVal tDB As String, ByVal SQLConnectionString As String, ByVal SQLuser As String, ByVal User400 As String, ByVal gGRID As DataGridView, ByVal tLOCX As String) As Integer
        Try
            Dim tMaxNo As String = SQL_READ_FIELD(gGRID, tDB, "MAXNO", SQLConnectionString, SQLuser, "SELECT MAX(RGMSG#) AS MAXNO FROM ROIDATA.RGMSGP WHERE RGLOCX='" + tLOCX + "'")
            Return Val(tMaxNo)
        Catch ex As Exception
            MSG_ERROR("NOTES_MAXNUMBER", ex.ToString)
        End Try
        Return False
    End Function

    Public Sub NOTES_PLUS(ByVal tDB As String, ByVal SQLConnectionString As String, ByVal SQLuser As String, ByVal User400 As String, ByVal gGRID As DataGridView, ByVal tLOCX As String, ByVal InOrOut As String)
        Try
            '**************************************************************************************************************
            '* 2011-08-23 RFK: The RACCTP RAMSGS must contain the EXACT number of NOTES in it, so it is displayed correctly
            '* 2012-01-12 RFK: moved this modified version to RKUTILS for better LIBRARY STANDARDIZATION
            Select Case tDB
                Case "DB2"
#If DB2 = 1 Then
                    Dim iMSGnumber As Integer = rkutils.NOTES_MAXNUMBER(tDB, SQLConnectionString, SQLuser, User400, gGRID, tLOCX)
                    Dim db2SQLCommandString As String = ""
                    If iMSGnumber > 0 Then
                        db2SQLCommandString = "UPDATE ROIDATA.RACCTP SET RAMSGS = " + iMSGnumber.ToString
                    Else
                        db2SQLCommandString = "UPDATE ROIDATA.RACCTP SET RAMSGS = 1"
                    End If
                    Select Case InOrOut
                        Case "O"
                            db2SQLCommandString = ", RAOUT# = RAOUT# + 1"
                    End Select
                    db2SQLCommandString += " WHERE RALOCX='" + tLOCX + "'"
                    rkutils.DB_COMMAND(tDB, SQLConnectionString, SQLuser, db2SQLCommandString)
#End If
                Case "MSSQL"
                    Dim msSQLCommandString As String = ""
            End Select
        Catch ex As Exception
            MSG_ERROR("NOTES_PLUS", ex.ToString)
        End Try
    End Sub

    Public Function LOCX_StatusMatchedAccounts(ByVal tDB As String, ByVal DB2SQLConnectionString As String, ByVal DB2SQLuser As String, ByVal MSSQLConnectionString As String, ByVal MSSQLuser As String, ByVal gGRID As DataGridView, ByVal tRamLOCX As String, ByVal tContactCode As String, ByVal tRAC As String, ByVal tSTATUS As String, ByVal tNOTE As String, ByVal TDate As String, ByVal TBy As String) As Boolean
        Try
            Dim tLocx As String = "", tSQL As String = "SELECT RALOCX,RARSTA FROM ROIDATA.RACCTP WHERE RAMLOCX='" + tRamLOCX + "'"
            SQL_READ_FIELD(gGRID, tDB, "RARSTA", DB2SQLConnectionString, DB2SQLuser, tSQL)
            For i1 = 0 To gGRID.RowCount - 1
                tLocx = DataGridView_ValueByColumnName(gGRID, "RALOCX", i1)
                If tLocx.Length > 0 Then
                    eProcess.MsgStatus("Matched Locx:" + tLocx, True)
                End If
            Next
            Return True
        Catch ex As Exception
            MSG_ERROR("LOCX_StatusMatchedAccounts", ex.ToString)
            Return False
        End Try
    End Function

    Public Function EMAILIT(ByVal tSQLConnection As String, ByVal tSQLuser As String, ByVal tEmailFrom As String, ByVal tEmailFromName As String, ByVal tEmailTo As String, ByVal tEmailToName As String, ByVal tModule As String, ByVal tSubject As String, ByVal tMessage As String, ByVal tHTML As String, ByVal tATTACH As String) As Boolean
        Try
            Dim SQLstring As String = "INSERT INTO RevMD.dbo.commands "
            SQLstring += " (COMMAND, TPARAMETERS, EMAILFROM, EMAILFROMNAME, EMAILTO, EMAILTONAME, EMAILSUBJECT, EMAILMESSAGE, EMAILATTACH, INSERT_DATE"
            SQLstring += ") VALUES("
            SQLstring += "'EMAIL'"
            SQLstring += ", ''"
            SQLstring += ", '" + STR_TRIM(STR_NORMALIZE(tEmailFrom), 100) + "'"
            SQLstring += ", '" + STR_TRIM(STR_NORMALIZE(tEmailFromName), 100) + "'"
            SQLstring += ", '" + STR_TRIM(STR_NORMALIZE(tEmailTo), 200) + "'"
            SQLstring += ", '" + STR_TRIM(STR_NORMALIZE(tEmailToName), 100) + "'"
            SQLstring += ", '" + STR_TRIM(STR_NORMALIZE(tSubject), 50) + "'"
            If tHTML.Contains("<html>") Then
                SQLstring += ", '" + STR_NORMALIZE(tHTML) + "'"
            Else
                SQLstring += ", '" + STR_NORMALIZE(tMessage) + "'"
            End If
            SQLstring += ", '" + STR_TRIM(STR_NORMALIZE(tATTACH), 200) + "'"
            SQLstring += ", '" + STR_format("TODAY", "mm/dd/ccyy HH:MM:SS") + "'"
            SQLstring += ")"
            rkutils.DB_COMMAND("MSSQL", tSQLConnection, tSQLuser, SQLstring)
            eProcess.MsgStatus(SQLstring, False)
            Return True
        Catch ex As Exception
            MSG_ERROR("EMAILIT", ex.ToString)
        End Try
        Return False
    End Function

    Public Function SQL_READ_DATAGRID(ByVal gGRID As DataGridView, ByVal tDB As String, ByVal tMODULE As String, ByVal tSQLConnectionString As String, ByVal tSQLuser As String, ByVal tSELECTstring As String) As Boolean
        Try
            Select Case tDB.ToUpper
                Case "DB2"
#If DB2 = 1 Then
                    Dim dbConnection As IBM.Data.DB2.iSeries.iDB2Connection = New IBM.Data.DB2.iSeries.iDB2Connection(tSQLConnectionString + tSQLuser)
                    Dim dbCommand As IBM.Data.DB2.iSeries.iDB2Command = New IBM.Data.DB2.iSeries.iDB2Command()
                    dbCommand.CommandText = tSELECTstring
                    dbCommand.Connection = dbConnection
                    dbCommand.CommandTimeout = 0
                    'dbCommand.ExecuteNonQuery()

                    Dim dataAdapter As New IBM.Data.DB2.iSeries.iDB2DataAdapter
                    dataAdapter.SelectCommand = dbCommand
                    Dim dataSet As System.Data.DataSet = New System.Data.DataSet
                    'dataAdapter.Fill(dataSet, "temp")

                    Dim i As Integer = dataAdapter.Fill(dataSet, "temp")
                    'Try dataAdapter.Fill(dataset, "temp")
                    'Catch ex7 As IBM.Data.DB2.iSeries.iDB2ExitProgramErrorException
                    '    MSG_ERROR("SQL_READ_DATAGRID/iDB2Exception", ex7.ToString)
                    '    Return False
                    'Catch eX3 As IBM.Data.DB2.iSeries.iDB2NullValueException
                    '    MSG_ERROR("SQL_READ_DATAGRID/iDB2SQLErrorException", eX3.ToString)
                    'Catch eX1 As IBM.Data.DB2.iSeries.iDB2TransactionFailedException
                    '    MSG_ERROR("SQL_READ_DATAGRID/iDB2TransactionFailedException", eX1.ToString)
                    '    Return False
                    'Catch eX2 As IBM.Data.DB2.iSeries.iDB2SQLErrorException
                    '    MSG_ERROR("SQL_READ_DATAGRID/iDB2SQLErrorException", eX2.ToString)
                    '    Return False
                    '    Return False
                    'Catch eX4 As IBM.Data.DB2.iSeries.iDB2SQLParameterErrorException
                    '    MSG_ERROR("SQL_READ_DATAGRID/iDB2SQLParameterErrorException", eX4.ToString)
                    '    Return False
                    'Catch ex6 As IBM.Data.DB2.iSeries.iDB2HostErrorException
                    '    MSG_ERROR("SQL_READ_DATAGRID/iDB2Exception", ex6.ToString)
                    '    Return False
                    'Catch eX5 As IBM.Data.DB2.iSeries.iDB2Exception
                    '    MSG_ERROR("SQL_READ_DATAGRID/iDB2Exception", eX5.ToString)
                    '    Return False
                    'Catch ex0 As Exception
                    '    MSG_ERROR("SQL_READ_DATAGRID", ex0.ToString)
                    '    Return False
                    'End Try


                    gGRID.DataSource = dataSet.Tables(0)
                    'gGRID.DataSource = DataTable
                    gGRID.Visible = False
                    dbCommand.Dispose()
                    dbCommand = Nothing
                    dbConnection.Close()
                    dbConnection.Dispose()
                    dbConnection = Nothing

                    If gGRID.Rows.Count > 0 Then
                        Return True
                    End If
#End If
                Case "MSSQL"
                    Dim dbConnection As New SqlConnection(tSQLConnectionString + tSQLuser)        ' The SqlConnection class allows you to communicate with SQL Server.
                    Dim dbCommand As New SqlCommand(tSELECTstring, dbConnection)                ' A SqlCommand object is used to execute the SQL commands.

                    Dim da As New SqlDataAdapter(dbCommand)
                    Dim mDataSet As New DataSet()
                    da.Fill(mDataSet, "temp")

                    gGRID.DataSource = mDataSet.Tables(0)
                    'gGRID.DataBind()
                    gGRID.Visible = False
                    dbCommand.Dispose()
                    dbCommand = Nothing
                    dbConnection.Close()
                    dbConnection.Dispose()
                    dbConnection = Nothing
                    If gGRID.Rows.Count > 0 Then
                        Return True
                    End If
                Case "MYSQL"
                    'Dim dbConnection As New MySqlConnection(tSQLConnectionString + tSQLuser)    'The SqlConnection class allows you to communicate with SQL Server.
                    'Dim dbCommand As New MySqlCommand(tSELECTstring, dbConnection)            'A SqlCommand object is used to execute the SQL commands.

                    'Dim da As New MySqlDataAdapter(dbCommand)
                    'Dim mDataSet As New DataSet()
                    'da.Fill(mDataSet, "temp")

                    'gGRID.DataSource = mDataSet.Tables(0)
                    ''gGRID.DataBind()
                    'gGRID.Visible = False
                    'dbCommand.Dispose()
                    'dbCommand = Nothing
                    'dbConnection.Close()
                    'dbConnection.Dispose()
                    'dbConnection = Nothing
                    'If gGRID.Rows.Count > 0 Then
                    '    Return True
                    'End If
                Case "FOXPRO"
                    Dim dbConnection As New OleDbConnection("Provider=vfpoledb.1;Data Source=" + tSQLConnectionString + ";Collating Sequence=machine")
                    Dim dbCommand As New OleDbCommand

                    Dim dbDataAdapter As New OleDbDataAdapter
                    Dim dbDataTable As New DataTable

                    dbCommand.CommandText = tSELECTstring
                    dbCommand.Connection = dbConnection

                    dbDataAdapter.SelectCommand = dbCommand
                    dbDataAdapter.Fill(dbDataTable)

                    gGRID.DataSource = dbDataTable
                    'gGRID.DataBind()
                    gGRID.Visible = False
                    dbCommand.Dispose()
                    dbCommand = Nothing
                    dbConnection.Close()
                    dbConnection.Dispose()
                    dbConnection = Nothing
                    If gGRID.Rows.Count > 0 Then
                        Return True
                    End If
            End Select
        Catch ex As Exception
            'MsgBox(ex.ToString)
            MSG_ERROR("SQL_READ_DATAGRID:" + tSELECTstring, ex.ToString)
        End Try
        Return False
    End Function

    Public Sub MSG_ERROR(ByVal tModule As String, ByVal tError As String)
        eProcess.MsgStatus(tModule + "/ERROR:" + tError, True)
        'EMAILIT()
    End Sub

    Public Sub ShellIT(ByVal sApp As String, ByVal sArguments As String, ByVal iFocus As Integer)
        Try
            eProcess.MsgStatus("ShellIT:" + sApp + " " + sArguments, True)
            If IS_File(sApp) Then
                Dim myProcess As New Process
                myProcess.StartInfo.UseShellExecute = True
                myProcess.StartInfo.FileName = sApp
                myProcess.StartInfo.Arguments = sArguments
                Select Case iFocus
                    Case 0
                        myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
                    Case 1
                        myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal
                    Case Else
                        myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal
                End Select
                myProcess.Start()
                Wait(2)     'Wait 2 seconds before going to the next one
            End If
        Catch ex As Exception
            MSG_ERROR("CCRUN", ex.ToString)
        End Try
    End Sub

    Public Function STR_format(ByVal tSTRin As String, ByVal tFormat As String) As String
        Try
            If tSTRin.ToUpper = "TODAY" And tFormat = "ISVALID" Then tSTRin = Now.Month.ToString.PadLeft(2, "0") + "-" + Now.Day.ToString.PadLeft(2, "0") + "-" + Now.Year.ToString
            If tSTRin = "TODAY" Then
                tSTRin = Now.Month.ToString.PadLeft(2, "0") + "/" + Now.Day.ToString.PadLeft(2, "0") + "/" + Now.Year.ToString
                If tFormat.Contains("HH") Or tFormat.Contains("MM") Or tFormat.Contains("SS") Then
                    tSTRin += " " + Now.Hour.ToString.PadLeft(2, "0") + ":" + Now.Minute.ToString.PadLeft(2, "0") + ":" + Now.Second.ToString.PadLeft(2, "0")
                End If
            End If
            If tSTRin.ToUpper = "YESTERDAY" Then tSTRin = DateAdd(DateInterval.Day, -1, Today).ToString
            If tSTRin.ToUpper = "TODAY-1" Then tSTRin = DateAdd(DateInterval.Day, -1, Today).ToString
            If tSTRin.ToUpper = "TODAY-2" Then tSTRin = DateAdd(DateInterval.Day, -2, Today).ToString
            If tSTRin.ToUpper = "TODAY-3" Then tSTRin = DateAdd(DateInterval.Day, -3, Today).ToString
            If tSTRin.ToUpper = "TODAY-4" Then tSTRin = DateAdd(DateInterval.Day, -4, Today).ToString
            If tSTRin.ToUpper = "TODAY-5" Then tSTRin = DateAdd(DateInterval.Day, -5, Today).ToString
            If tSTRin.ToUpper = "TODAY-6" Then tSTRin = DateAdd(DateInterval.Day, -6, Today).ToString
            If tSTRin.ToUpper = "TODAY-7" Then tSTRin = DateAdd(DateInterval.Day, -7, Today).ToString
            '****************************
            Dim tSTRout As String = tSTRin
            Select Case tFormat
                Case "%"
                    tSTRout = String.Format("{0:#.##}", Val(tSTRin))
                Case "ISVALID", "VALID"
                    tSTRout = ""
                    Dim iVAL As Integer
                    For i1 = 0 To tSTRin.Length - 1
                        iVAL = Asc(tSTRin.Substring(i1, 1))
                        Select Case iVAL
                            Case Asc("a") To Asc("z")
                                tSTRout += tSTRin.Substring(i1, 1)
                            Case Asc("A") To Asc("Z")
                                tSTRout += tSTRin.Substring(i1, 1)
                            Case Asc("0") To Asc("9")
                                tSTRout += tSTRin.Substring(i1, 1)
                            Case Asc(" "), Asc("_"), Asc("#"), Asc("$"), Asc("%"), Asc("&"), Asc("("), Asc(")"), Asc("-"), Asc("+"), Asc("."), Asc("?"), Asc("<"), Asc(">"), Asc("="), Asc("@"), Asc(":"), Asc(","), Asc("["), Asc("]")
                                tSTRout += tSTRin.Substring(i1, 1)
                            Case Else
                                If tFormat = "ISVALID" Then
                                    Return ""
                                End If
                                tSTRout += "_"
                        End Select
                    Next
                Case "$"
                    tSTRout = String.Format("{0:C}", Val(tSTRin))
                Case "#"
                    tSTRout = String.Format("{0:#.##}", Val(tSTRin))
                Case "0"
                    tSTRout = String.Format("{0:0.00}", Val(tSTRin))
                Case "PHONE"
                    If tSTRin.Length = 10 Then
                        tSTRout = tSTRin.Substring(0, 3) + "-" + tSTRin.Substring(3, 3) + "-" + tSTRin.Substring(6, 4)
                    End If
                Case "SSN"
                    If tSTRin.Length = 9 Then
                        tSTRout = tSTRin.Substring(0, 3) + "-" + tSTRin.Substring(3, 2) + "-" + tSTRin.Substring(5, 4)
                    End If
                Case "DOW"
                    If IsDate(tSTRin) Then
                        Dim dDATE As Date = tSTRin
                        Select Case dDATE.DayOfWeek
                            Case 0
                                tSTRout = "Sun"
                            Case 1
                                tSTRout = "Mon"
                            Case 2
                                tSTRout = "Tue"
                            Case 3
                                tSTRout = "Wed"
                            Case 4
                                tSTRout = "Thu"
                            Case 5
                                tSTRout = "Fri"
                            Case 6
                                tSTRout = "Sat"
                        End Select
                    End If
                Case "DAYOFWEEK"
                    If IsDate(tSTRin) Then
                        Dim dDATE As Date = tSTRin
                        Select Case dDATE.DayOfWeek
                            Case 0
                                tSTRout = "SUNDAY"
                            Case 1
                                tSTRout = "MONDAY"
                            Case 2
                                tSTRout = "TUESDAY"
                            Case 3
                                tSTRout = "WEDNESDAY"
                            Case 4
                                tSTRout = "THURSDAY"
                            Case 5
                                tSTRout = "FRIDAY"
                            Case 6
                                tSTRout = "SATURDAY"
                        End Select
                    End If
                Case "ccyymmdd"
                    If tSTRin = "01011900" Then
                        tSTRout = ""
                    Else
                        Dim dDATE As Date = tSTRin
                        tSTRout = dDATE.Year.ToString
                        If dDATE.Month >= 10 Then
                            tSTRout += dDATE.Month.ToString
                        Else
                            tSTRout += "0" + dDATE.Month.ToString
                        End If
                        If dDATE.Day >= 10 Then
                            tSTRout += dDATE.Day.ToString
                        Else
                            tSTRout += "0" + dDATE.Day.ToString
                        End If
                    End If
                Case "ccyy-mm-dd"
                    If tSTRin = "01011900" Then
                        tSTRout = ""
                    Else
                        Dim dDATE As Date = tSTRin
                        tSTRout = dDATE.Year.ToString
                        tSTRout += "-"
                        If dDATE.Month >= 10 Then
                            tSTRout += dDATE.Month.ToString
                        Else
                            tSTRout += "0" + dDATE.Month.ToString
                        End If
                        tSTRout += "-"
                        If dDATE.Day >= 10 Then
                            tSTRout += dDATE.Day.ToString
                        Else
                            tSTRout += "0" + dDATE.Day.ToString
                        End If
                    End If
                Case "mm/dd/ccyy"
                    If tSTRin.Trim.Length >= 8 Then
                        If tSTRin.Substring(0, 8) = "1/1/1800" Then 'Or tSTRin.Substring(0, 8) = "1/1/1900" Then
                            tSTRout = ""
                        Else
                            If tSTRin.Contains("/") = False And tSTRin.Substring(0, 2) = "19" Or tSTRin.Substring(0, 2) = "20" Then
                                tSTRin = tSTRin.Substring(4, 2) + "/" + tSTRin.Substring(6, 2) + "/" + tSTRin.Substring(0, 4)
                            End If
                            Dim dDATE As Date = tSTRin
                            If dDATE.Month >= 10 Then
                                tSTRout = dDATE.Month.ToString
                            Else
                                tSTRout = "0" + dDATE.Month.ToString
                            End If
                            tSTRout += "/"
                            If dDATE.Day >= 10 Then
                                tSTRout += dDATE.Day.ToString
                            Else
                                tSTRout += "0" + dDATE.Day.ToString
                            End If
                            tSTRout += "/"
                            tSTRout += dDATE.Year.ToString
                        End If
                    End If
                Case "mmdd"
                    If IsDate(tSTRin) Then
                        Dim dDATE As Date = tSTRin
                        tSTRout = dDATE.Month.ToString.PadLeft(2, "0")
                        tSTRout += dDATE.Day.ToString.PadLeft(2, "0")
                    End If
                Case "mm-dd-ccyy"
                    If IsDate(tSTRin) Then
                        Dim dDATE As Date = tSTRin
                        tSTRout = dDATE.Month.ToString.PadLeft(2, "0")
                        tSTRout += "-"
                        tSTRout += dDATE.Day.ToString.PadLeft(2, "0")
                        tSTRout += "-"
                        tSTRout += dDATE.Year.ToString.PadLeft(4, "0")
                    End If
                Case "mm"
                    If tSTRin.Trim.Length >= 8 Then
                        If tSTRin.Substring(0, 8) = "1/1/1800" Then
                            tSTRout = ""
                        Else
                            If tSTRin.Contains("/") = False And tSTRin.Substring(0, 2) = "19" Or tSTRin.Substring(0, 2) = "20" Then
                                tSTRin = tSTRin.Substring(4, 2) + "/" + tSTRin.Substring(6, 2) + "/" + tSTRin.Substring(0, 4)
                            End If
                            Dim dDATE As Date = tSTRin
                            tSTRout = dDATE.Month.ToString.PadLeft(2, "0")
                        End If
                    End If
                Case "dd"
                    If tSTRin.Trim.Length >= 8 Then
                        If tSTRin.Substring(0, 8) = "1/1/1800" Then 'Or tSTRin.Substring(0, 8) = "1/1/1900" Then
                            tSTRout = ""
                        Else
                            If tSTRin.Contains("/") = False And tSTRin.Substring(0, 2) = "19" Or tSTRin.Substring(0, 2) = "20" Then
                                tSTRin = tSTRin.Substring(4, 2) + "/" + tSTRin.Substring(6, 2) + "/" + tSTRin.Substring(0, 4)
                            End If
                            Dim dDATE As Date = tSTRin
                            tSTRout = dDATE.Day.ToString.PadLeft(2, "0")
                        End If
                    End If
                Case "ccyy"
                    If tSTRin.Trim.Length >= 8 Then
                        If tSTRin.Substring(0, 8) = "1/1/1800" Then 'Or tSTRin.Substring(0, 8) = "1/1/1900" Then
                            tSTRout = ""
                        Else
                            If tSTRin.Contains("/") = False And tSTRin.Substring(0, 2) = "19" Or tSTRin.Substring(0, 2) = "20" Then
                                tSTRin = tSTRin.Substring(4, 2) + "/" + tSTRin.Substring(6, 2) + "/" + tSTRin.Substring(0, 4)
                            End If
                            Dim dDATE As Date = tSTRin
                            tSTRout = dDATE.Year.ToString
                        End If
                    End If
                Case "yy"
                    If tSTRin.Trim.Length >= 8 Then
                        If tSTRin.Substring(0, 8) = "1/1/1800" Then 'Or tSTRin.Substring(0, 8) = "1/1/1900" Then
                            tSTRout = ""
                        Else
                            If tSTRin.Contains("/") = False And tSTRin.Substring(0, 2) = "19" Or tSTRin.Substring(0, 2) = "20" Then
                                tSTRin = tSTRin.Substring(4, 2) + "/" + tSTRin.Substring(6, 2) + "/" + tSTRin.Substring(0, 4)
                            End If
                            Dim dDATE As Date = tSTRin
                            tSTRout = Right(dDATE.Year.ToString, 2)
                        End If
                    End If
                Case "ccyy/mm/dd"
                    If tSTRin = "01011900" Then
                        tSTRout = ""
                    Else
                        If IsDate(tSTRin) Then
                            Dim dDATE As Date = tSTRin
                            tSTRout = dDATE.Year.ToString
                            tSTRout += "/"
                            If dDATE.Month >= 10 Then
                                tSTRout += dDATE.Month.ToString
                            Else
                                tSTRout += "0" + dDATE.Month.ToString
                            End If
                            tSTRout += "/"
                            If dDATE.Day >= 10 Then
                                tSTRout += dDATE.Day.ToString
                            Else
                                tSTRout += "0" + dDATE.Day.ToString
                            End If
                        End If
                    End If
                Case "mmddccyy"
                    If tSTRin.Length >= 8 Then
                        If IsDate(tSTRin) Then
                            If tSTRin.Substring(0, 8) = "1/1/1800" Then 'Or tSTRin.Substring(0, 8) = "1/1/1900" Then
                                tSTRout = ""
                            Else
                                Dim dDATE As Date = tSTRin
                                If dDATE.Month >= 10 Then
                                    tSTRout = dDATE.Month.ToString
                                Else
                                    tSTRout = "0" + dDATE.Month.ToString
                                End If
                                If dDATE.Day >= 10 Then
                                    tSTRout += dDATE.Day.ToString
                                Else
                                    tSTRout += "0" + dDATE.Day.ToString
                                End If
                                tSTRout += dDATE.Year.ToString
                            End If
                        End If
                    End If
                Case "ccyymmddHHMMSSss", "ccyymmddHHMMSS", "ccyymmdd_HHMMSS", "ccyymmdd_HHMM", "ccyymmdd_HHMMSSss"
                    Dim dDATE As Date = tSTRin
                    tSTRout = dDATE.Year.ToString
                    If dDATE.Month >= 10 Then
                        tSTRout += dDATE.Month.ToString
                    Else
                        tSTRout += "0" + dDATE.Month.ToString
                    End If
                    If dDATE.Day >= 10 Then
                        tSTRout += dDATE.Day.ToString
                    Else
                        tSTRout += "0" + dDATE.Day.ToString
                    End If
                    '**********************
                    '* 2012-02-27 RFK: 
                    If tFormat = "ccyymmdd_HHMM" Or tFormat = "ccyymmdd_HHMMSS" Or tFormat = "ccyymmdd_HHMMSSss" Then
                        tSTRout += "_"
                    End If
                    '**********************
                    '* 2012-02-27 RFK: 
                    If dDATE.Hour = 0 Then
                        If Now.Hour >= 10 Then
                            tSTRout += Now.Hour.ToString
                        Else
                            tSTRout += "0" + Now.Hour.ToString
                        End If
                        If Now.Minute >= 10 Then
                            tSTRout += Now.Minute.ToString
                        Else
                            tSTRout += "0" + Now.Minute.ToString
                        End If
                        If Now.Second >= 10 Then
                            tSTRout += Now.Second.ToString
                        Else
                            tSTRout += "0" + Now.Second.ToString
                        End If
                        If tFormat = "ccyymmddHHMMSSss" Or tFormat = "ccyymmdd_HHMMSSss" Then
                            If Now.Millisecond >= 10 Then
                                tSTRout += Now.Millisecond.ToString
                            Else
                                tSTRout += "0" + Now.Millisecond.ToString
                            End If
                        End If
                    Else
                        If dDATE.Hour >= 10 Then
                            tSTRout += dDATE.Hour.ToString
                        Else
                            tSTRout += "0" + dDATE.Hour.ToString
                        End If
                        If dDATE.Minute >= 10 Then
                            tSTRout += dDATE.Minute.ToString
                        Else
                            tSTRout += "0" + dDATE.Minute.ToString
                        End If
                        If dDATE.Second >= 10 Then
                            tSTRout += dDATE.Second.ToString
                        Else
                            tSTRout += "0" + dDATE.Second.ToString
                        End If
                        If tFormat = "ccyymmddHHMMSSss" Or tFormat = "ccyymmdd_HHMMSSss" Then
                            If dDATE.Millisecond >= 10 Then
                                tSTRout += dDATE.Millisecond.ToString
                            Else
                                tSTRout += "0" + dDATE.Millisecond.ToString
                            End If
                        End If
                    End If

                    If tSTRout.Length > 16 Then
                        tSTRout = tSTRout.Substring(0, 16)
                    End If
                Case "ccyy-mm-dd HH:MM:SS"
                    Dim dDATE As Date = tSTRin
                    tSTRout = dDATE.Year.ToString
                    tSTRout += "-"
                    If dDATE.Month >= 10 Then
                        tSTRout += dDATE.Month.ToString
                    Else
                        tSTRout += "0" + dDATE.Month.ToString
                    End If
                    tSTRout += "-"
                    If dDATE.Day >= 10 Then
                        tSTRout += dDATE.Day.ToString
                    Else
                        tSTRout += "0" + dDATE.Day.ToString
                    End If
                    tSTRout += " "
                    If dDATE.Hour >= 10 Then
                        tSTRout += dDATE.Hour.ToString
                    Else
                        tSTRout += "0" + dDATE.Hour.ToString
                    End If
                    tSTRout += ":"
                    If dDATE.Minute >= 10 Then
                        tSTRout += dDATE.Minute.ToString
                    Else
                        tSTRout += "0" + dDATE.Minute.ToString
                    End If
                    tSTRout += ":"
                    If dDATE.Second >= 10 Then
                        tSTRout += dDATE.Second.ToString
                    Else
                        tSTRout += "0" + dDATE.Second.ToString
                    End If
                Case "HH"
                    If IsDate(tSTRin) Then
                        Dim dDATE As Date = tSTRin
                        tSTRout = dDATE.Hour.ToString.PadLeft(2, "0")
                    End If
                Case "HHMM"
                    If IsDate(tSTRin) Then
                        Dim dDATE As Date = tSTRin
                        tSTRout = dDATE.Hour.ToString.PadLeft(2, "0")
                        tSTRout += dDATE.Minute.ToString.PadLeft(2, "0")
                    End If
                Case "HHMMSS"
                    If tSTRin.Length = 8 And tSTRin.Contains(":") Then
                        tSTRout = tSTRin.Replace(":", "")
                    Else
                        'eProcess.MsgStatus("CHECK:" + tSTRin, True)
                        If IsDate(tSTRin) Then
                            Dim dDATE As Date = tSTRin
                            'eProcess.MsgStatus("Hour:" + dDATE.Hour.ToString + " Min:" + dDATE.Minute.ToString + " Sec:" + dDATE.Second.ToString, True)
                            tSTRout = dDATE.Hour.ToString.PadLeft(2, "0")
                            tSTRout += dDATE.Minute.ToString.PadLeft(2, "0")
                            tSTRout += dDATE.Second.ToString.PadLeft(2, "0")
                            'eProcess.MsgStatus(tSTRout, True)
                        End If
                    End If
                Case "HH:MM"
                    If IsDate(tSTRin) Then
                        Dim dDATE As Date = tSTRin
                        tSTRout = dDATE.Hour.ToString.PadLeft(2, "0")
                        tSTRout += ":"
                        tSTRout += dDATE.Minute.ToString.PadLeft(2, "0")
                    End If
                Case "HH:MM:SS"
                    If IsDate(tSTRin) Then
                        Dim dDATE As Date = tSTRin
                        tSTRout = dDATE.Hour.ToString.PadLeft(2, "0")
                        tSTRout += ":"
                        tSTRout += dDATE.Minute.ToString.PadLeft(2, "0")
                        tSTRout += ":"
                        tSTRout += dDATE.Second.ToString.PadLeft(2, "0")
                    End If
                Case "HH:MM:SS."
                    If IsDate(tSTRin) Then
                        Dim dDATE As Date = tSTRin
                        tSTRout = dDATE.Hour.ToString.PadLeft(2, "0")
                        tSTRout += ":"
                        tSTRout += dDATE.Minute.ToString.PadLeft(2, "0")
                        tSTRout += ":"
                        tSTRout += dDATE.Second.ToString.PadLeft(2, "0")
                        tSTRout += "."
                        tSTRout += dDATE.Millisecond.ToString
                    End If
                Case "AGE_D"
                    If tSTRin.Length >= 8 Then
                        If tSTRin.Substring(0, 8) = "1/1/1800" Then
                            tSTRout = ""
                        Else
                            Dim dDATE As Date = tSTRin
                            tSTRout = Now.Date.Subtract(dDATE).Days.ToString
                        End If
                    Else
                        tSTRout = ""
                    End If
                Case "AGE_Y"
                    If tSTRin.Length >= 8 Then
                        If tSTRin.Substring(0, 8) = "1/1/1800" Then
                            tSTRout = ""
                        Else
                            Dim dDATE As Date = tSTRin
                            tSTRout = Now.Date.Subtract(dDATE).Days.ToString
                            If Val(tSTRout) > 365 Then
                                tSTRout = (Val(tSTRout) / 365).ToString.Trim
                                If tSTRout.Contains(".") Then
                                    tSTRout = tSTRout.Substring(0, tSTRout.IndexOf("."))
                                End If
                            End If
                        End If
                    Else
                        tSTRout = ""
                    End If
            End Select
            Return tSTRout
        Catch ex As Exception
            MSG_ERROR("STR_format:" + tSTRin + " " + tFormat, ex.ToString)
        End Try
        Return ""
    End Function

    Public Sub ExcelFromDataGrid(ByVal sFileName As String, ByVal DGV As DataGridView, ByVal bOpenDuring As Boolean, ByVal bLeaveOpen As Boolean)
        Try
            '**********************************************************************
            '* 2015-05-22 RFK:
            Dim excelApp As Object, excelBook As Object, excelSheet As Object
            excelApp = CreateObject("Excel.Application")
            excelBook = excelApp.workbooks.add
            excelSheet = excelBook.worksheets(1)
            excelApp.Visible = bOpenDuring
            '**********************************************************************
            For Each column As DataGridViewColumn In DGV.Columns
                excelSheet.cells(1, column.Index + 1) = column.HeaderText
            Next
            For i = 1 To DGV.RowCount - 1
                Application.DoEvents()
                eProcess.Label_AnHour.Text = Trim(Str(Val(eProcess.Label_AnHour.Text) + 1))
                '**********************************************************************
                For j = 0 To DGV.Columns.Count - 1
                    If DGV.Rows(i - 1).Cells(j).Value IsNot Nothing Then
                        excelSheet.cells(i, j + 1) = DGV.Rows(i - 1).Cells(j).Value
                    End If
                Next
            Next
            '**********************************************************************
            excelApp.Visible = bLeaveOpen
            excelSheet = Nothing
            excelBook = Nothing
            excelApp = Nothing
            '**********************************************************************
        Catch ex As Exception

        End Try
    End Sub

    Public Sub ExportToExcel(ByVal sFileName As String, ByVal DGV As DataGridView)
        Try
            Dim fs As New StreamWriter(sFileName, False)
            With fs
                .WriteLine("<?xml version=""1.0""?>")
                .WriteLine("<?mso-application progid=""Excel.Sheet""?>")
                .WriteLine("<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"">")
                .WriteLine("    <Styles>")
                .WriteLine("        <Style ss:ID=""hdr"">")
                .WriteLine("            <Alignment ss:Horizontal=""Center""/>")
                .WriteLine("            <Borders>")
                .WriteLine("                <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>")
                .WriteLine("                <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>")
                .WriteLine("                <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>")
                .WriteLine("            </Borders>")
                .WriteLine("            <Font ss:FontName=""Calibri"" ss:Size=""11"" ss:Bold=""1""/>") 'SET FONT
                .WriteLine("        </Style>")
                .WriteLine("        <Style ss:ID=""ksg"">")
                .WriteLine("            <Alignment ss:Vertical=""Bottom""/>")
                .WriteLine("            <Borders/>")
                .WriteLine("            <Font ss:FontName=""Calibri""/>") 'SET FONT
                .WriteLine("        </Style>")
                .WriteLine("        <Style ss:ID=""isi"">")
                .WriteLine("            <Borders>")
                .WriteLine("                <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>")
                .WriteLine("                <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>")
                .WriteLine("                <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>")
                .WriteLine("                <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>")
                .WriteLine("            </Borders>")
                .WriteLine("            <Font ss:FontName=""Calibri"" ss:Size=""10""/>") 'SET FONT
                .WriteLine("        </Style>")
                .WriteLine("    </Styles>")
                If DGV.Name = "Student" Then
                    .WriteLine("    <Worksheet ss:Name=""Student"">") 'SET NAMA SHEET
                    .WriteLine("        <Table>")
                    .WriteLine("            <Column ss:Width=""27.75""/>") 'No
                    .WriteLine("            <Column ss:Width=""93""/>") 'NIK
                    .WriteLine("            <Column ss:Width=""84""/>") 'Nama
                    .WriteLine("            <Column ss:Width=""100""/>") 'Alamat
                    .WriteLine("            <Column ss:Width=""84""/>") 'Telp
                End If
                'AUTO SET HEADER
                .WriteLine("            <Row ss:StyleID=""ksg"">")
                For i As Integer = 0 To DGV.Columns.Count - 1 'SET HEADER
                    Application.DoEvents()
                    If DGV.Columns.Item(i).HeaderText IsNot Nothing Then
                        .WriteLine("            <Cell ss:StyleID=""hdr"">")
                        .WriteLine("                <Data ss:Type=""String"">{0}</Data>", DGV.Columns.Item(i).HeaderText)
                        .WriteLine("            </Cell>")
                    End If
                Next
                .WriteLine("            </Row>")
                eProcess.Label_AnHour.Text = "0"
                For intRow As Integer = 0 To DGV.RowCount - 1
                    Application.DoEvents()
                    eProcess.Label_AnHour.Text = Trim(Str(Val(eProcess.Label_AnHour.Text) + 1))
                    .WriteLine("        <Row ss:StyleID=""ksg"" ss:utoFitHeight =""0"">")
                    For intCol As Integer = 0 To DGV.Columns.Count - 1
                        Application.DoEvents()
                        If DGV.Item(intCol, intRow).Value IsNot Nothing Then
                            .WriteLine("        <Cell ss:StyleID=""isi"">")
                            .WriteLine("            <Data ss:Type=""String"">{0}</Data>", DGV.Item(intCol, intRow).Value.ToString)
                            .WriteLine("        </Cell>")
                        End If
                    Next
                    .WriteLine("        </Row>")
                Next
                .WriteLine("        </Table>")
                .WriteLine("    </Worksheet>")
                .WriteLine("</Workbook>")
                .Close()
                eProcess.MsgStatus(sFileName + " created", True)
            End With
        Catch ex As Exception
            MSG_ERROR("ExportToExcel", ex.ToString)
        End Try
    End Sub

    Public Sub ExportToCSV(ByVal sFileName As String, ByVal sDelimiter As String, ByVal DGV As DataGridView)
        Try
            eProcess.MsgStatus(sFileName, True)
            Dim sTempStr As String = ""
            Dim fs As New StreamWriter(sFileName, False)
            With fs
                For i As Integer = 0 To DGV.Columns.Count - 1 'SET HEADER
                    Application.DoEvents()
                    If DGV.Columns.Item(i).HeaderText IsNot Nothing Then
                        sTempStr = Chr(34) + DGV.Columns.Item(i).HeaderText + Chr(34) + sDelimiter
                        .Write("{0}", sTempStr)
                    End If
                Next
                .WriteLine("")
                eProcess.Label_AnHour.Text = "0"
                For intRow As Integer = 0 To DGV.RowCount - 1
                    Application.DoEvents()
                    eProcess.Label_AnHour.Text = Trim(Str(Val(eProcess.Label_AnHour.Text) + 1))
                    For intCol As Integer = 0 To DGV.Columns.Count - 1
                        Application.DoEvents()
                        If DGV.Item(intCol, intRow).Value IsNot Nothing Then
                            sTempStr = Chr(34) + DGV.Item(intCol, intRow).Value.ToString + Chr(34) + sDelimiter
                            .Write("{0}", sTempStr)
                        End If
                    Next
                    .WriteLine("")
                Next
                .Close()
                eProcess.MsgStatus(sFileName + " created", True)
            End With
        Catch ex As Exception
            MSG_ERROR("ExportToExcel", ex.ToString)
        End Try
    End Sub

    Public Function DataTable_ColumnByName(ByVal DT As DataTable, ByVal tColName As String) As Integer
        Try
            '******************************************************************
            ' 2015-08-04 RFK: all the rows
            Dim i2 As Integer = 0
            For Each col As DataColumn In DT.Columns
                If col.ColumnName.ToUpper = tColName.ToUpper Then Return i2
                i2 += 1
            Next
        Catch ex As Exception
            'MSG_warning(ex.ToString)
        End Try
        Return -1
    End Function

    Public Function DataTable_ValueByColumnName(ByVal DT As DataTable, ByVal tColName As String, ByVal iRow As Integer) As String
        Try
            Dim iRC As Integer = DataTable_ColumnByName(DT, tColName)
            If iRC >= 0 Then
                If DT.Rows(iRow).Item(iRC) IsNot Nothing Then
                    If DT.Rows(iRow).Item(iRC).ToString.Trim.Length > 0 Then
                        Return DT.Rows(iRow).Item(iRC).ToString.Trim
                    End If
                End If
            End If
        Catch ex As Exception
            'MSG_warning(ex.ToString)
        End Try
        Return ""
    End Function

    Public Function ReadFieldDTable(ByVal DT As DataTable, ByVal tColumnName As String, ByVal iRow As Integer) As String
        Return rkutils.STR_convert_AMP(rkutils.STR_NORMALIZE(DataTable_ValueByColumnName(DT, tColumnName, iRow).Trim))
    End Function

    Public Function ReadFieldDateSelectStringDTable(ByVal DT As DataTable, ByVal tColumnName As String, ByVal iRow As Integer, ByVal FieldNameOrValue As Boolean, ByVal tFieldName As String) As String
        Try
            Dim sTemp As String = DataTable_ValueByColumnName(DT, tColumnName, iRow)
            If IsDate(sTemp) Then
                Dim dDate As Date = sTemp
                If dDate.Year > 1900 Then
                    If FieldNameOrValue Then
                        Return "," + tFieldName
                    Else
                        Return ",'" + STR_format(sTemp, "mm/dd/ccyy") + "'"
                    End If
                End If
            End If
            Return ""
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function ReadFieldDateCCYYMMDDSelectStringDTable(ByVal DT As DataTable, ByVal tColumnName As String, ByVal iRow As Integer, ByVal FieldNameOrValue As Boolean, ByVal tFieldName As String) As String
        Try
            Dim sTemp As String = DataTable_ValueByColumnName(DT, tColumnName, iRow).Trim
            Dim sTemp2 As String = ""
            If sTemp.Length = 8 Then
                sTemp2 = sTemp.Substring(4, 2) + "/" + sTemp.Substring(6, 2) + "/" + sTemp.Substring(0, 4)
                If IsDate(sTemp2) Then
                    Dim dDate As Date = sTemp2
                    If dDate.Year > 1900 Then
                        If FieldNameOrValue Then
                            Return "," + tFieldName
                        Else
                            Return ",'" + STR_format(sTemp2, "mm/dd/ccyy") + "'"
                        End If
                    End If
                End If
            End If
            Return ""
        Catch ex As Exception
            MSG_ERROR("ReadFieldDateCCYYMMDDSelectString", ex.ToString)
        End Try
        Return ""
    End Function

    Public Function STR_trimSQL(ByVal tSTR As String, ByVal iMaxChars As Integer) As String
        Try
            '******************************************************************
            '* 2016-04-29 RFK:
            Dim tNew As String = tSTR
            If tNew.Length > iMaxChars Then
                Return tNew.Substring(0, iMaxChars).Trim
            Else
                Return tNew
            End If
            Return ""
        Catch ex As Exception
            MSG_ERROR("STR_trim", ex.ToString)
        End Try
        Return ""
        '**********************************************************************
    End Function

    Public Function WhereAnd(ByVal tStrIN As String, ByVal tAdd As String) As String
        If tStrIN.Contains("WHERE") Then Return " AND " + tAdd
        Return " WHERE " + tAdd
    End Function

    Public Function WhereOr(ByVal tStrIN As String, ByVal tAdd As String) As String
        If tStrIN.Contains("WHERE") Then Return " OR " + tAdd
        Return " WHERE (" + tAdd
    End Function

    Public Function WhereOrClosing(ByVal tStrIN As String) As String
        If tStrIN.Contains("WHERE") And tStrIN.Contains("(") Then Return ")"
        Return ""
    End Function

    Public Function DataTable_to_CSV(ByVal dT As DataTable, ByVal sFileName As String, ByVal IncludeColumnHeader As Boolean, ByVal lLabelRowUpdate As Label) As Boolean
        Try
            '******************************************************************
            '* 2015-07-30 RFK: Check for existing
            Dim sep As String = ""
            Dim builder As New System.Text.StringBuilder
            If File.Exists(sFileName) Then
                eProcess.MsgStatus("Deleting:" + sFileName, True)
                File.Delete(sFileName)
            End If
            If File.Exists(sFileName) Then
                eProcess.MsgStatus("Unable to create, exists:" + sFileName, True)
                Return False
            End If
            '******************************************************************
            '* 2015-07-30 RFK: Open
            eProcess.MsgStatus("Creating:" + sFileName, True)
            Dim sw As System.IO.StreamWriter
            sw = My.Computer.FileSystem.OpenTextFileWriter(sFileName, True)
            '******************************************************************
            ' 2015-07-30 RFK: columns
            If IncludeColumnHeader Then
                sep = ""
                For Each col As DataColumn In dT.Columns
                    builder.Append(sep).Append(Trim(col.ColumnName))
                    sep = ","   'After 1st one now add the seperator
                Next
                sw.WriteLine(builder.ToString())
            End If
            '******************************************************************
            ' 2015-07-30 RFK: all the rows
            Dim iCTR As Integer = 0
            For Each row As DataRow In dT.Rows
                lLabelRowUpdate.Text = Val(lLabelRowUpdate.Text) - 1.ToString.Trim
                System.Windows.Forms.Application.DoEvents()
                '**************************************************************
                sep = ""
                builder = New System.Text.StringBuilder
                '**************************************************************
                iCTR = 0
                For Each col As DataColumn In dT.Columns
                    If row.Item(iCTR).ToString.Length > 0 Then
                        If row(col.ColumnName).ToString IsNot Nothing Then
                            builder.Append(sep).Append(Trim(row(col.ColumnName)).Replace(",", " "))
                        End If
                    End If
                    sep = ","   'After 1st one now add the seperator
                    iCTR += 1
                Next
                sw.WriteLine(builder.ToString())
            Next
            '******************************************************************
            '* 2015-07-30 RFK: 
            If Not sw Is Nothing Then sw.Close()
            lLabelRowUpdate.Text = "0"
            eProcess.MsgStatus("Wrote " + dT.Rows.ToString.Trim + " rows.", True)
            Return True
        Catch ex As Exception
            MSG_ERROR("DataTable_to_CSV", ex.ToString)
        End Try
        Return False
    End Function

    Public Function DataGridview_ToHTMLtable(ByVal gGrid As DataGridView, ByVal sType As String) As String
        Try
            Dim sHTML As String = ""
            Dim iRow As Integer = 0, iCol As Integer = 0
            '******************************************************************
            '* 2017-01-11 RFK: 
            sHTML = "<table width=100% cellpadding=2 cellspacing=2 border=1>"
            '******************************************************************
            sHTML += "<tr>"
            For iCol = 0 To gGrid.ColumnCount - 1
                sHTML += "<td>"
                If gGrid.Columns(iCol).Name IsNot Nothing Then sHTML += gGrid.Columns(iCol).Name
                sHTML += "</td>"
            Next
            sHTML += "</tr>"
            '******************************************************************
            For iRow = 0 To gGrid.RowCount - 2
                'rkutils.DoEvents()
                sHTML += "<tr>"
                For iCol = 0 To gGrid.ColumnCount - 1
                    '**********************************************************
                    sHTML += "<td>"
                    If gGrid.Item(iCol, iRow).Value.ToString() IsNot Nothing Then sHTML += gGrid.Item(iCol, iRow).Value.ToString
                    sHTML += "</td>"
                    '**********************************************************
                Next
                sHTML += "</tr>"
                '**************************************************************
            Next
            Return sHTML
        Catch ex As Exception
            'eProcess.MsgStatus(iRow.ToString, True)
            'eProcess.MsgStatus(sHTML, True)
            'eProcess.MsgStatus(ex.ToString, True)
            MSG_ERROR("DataTable_to_CSV", ex.ToString)
        End Try
        Return ""
    End Function

    Public Function WhatIsClientVMBValue(ByVal gTempGrid As DataGridView, ByVal tSQLConnectionString As String, ByVal tSQLuser As String, ByVal tClientName As String, ByVal tClientTOB As String, ByVal tClientFACILITY As String, ByVal tField As String) As String
        Try
            Dim tReturn As String = "", msSQLCommandString As String = ""
            '***************************************************
            '* 2012-04-06 RFK: only look for Client/TOB/facility
            If tClientName.Length > 0 And tClientTOB.Length > 0 And tClientFACILITY.Length > 0 Then
                msSQLCommandString = "SELECT * FROM RevMD.dbo.clientsVMB WHERE ClientName='" + tClientName + "' AND tob='" + tClientTOB + "' AND facility='" + tClientFACILITY + "'"
                If rkutils.SQL_READ_DATAGRID(gTempGrid, "MSSQL", "*", tSQLConnectionString, tSQLuser, msSQLCommandString) Then
                    tReturn = rkutils.DataGridView_ValueByColumnName(gTempGrid, tField, 0)
                    If tReturn.Length > 0 Then Return tReturn
                End If
            End If
            '***********************************************
            '* 2012-04-06 RFK: only look for Client/facility
            If tClientName.Length > 0 And tClientFACILITY.Length > 0 Then
                msSQLCommandString = "SELECT * FROM RevMD.dbo.clientsVMB WHERE ClientName='" + tClientName + "' AND tob='0' AND facility='" + tClientFACILITY + "'"
                If rkutils.SQL_READ_DATAGRID(gTempGrid, "MSSQL", "*", tSQLConnectionString, tSQLuser, msSQLCommandString) Then
                    tReturn = rkutils.DataGridView_ValueByColumnName(gTempGrid, tField, 0)
                    If tReturn.Length > 0 Then Return tReturn
                End If
            End If
            '**********************************************
            '* 2012-04-06 RFK: only look for Client
            If tClientName.Length > 0 Then
                msSQLCommandString = "SELECT * FROM RevMD.dbo.clientsVMB WHERE ClientName='" + tClientName + "'"
                If rkutils.SQL_READ_DATAGRID(gTempGrid, "MSSQL", "*", tSQLConnectionString, tSQLuser, msSQLCommandString) Then
                    tReturn = rkutils.DataGridView_ValueByColumnName(gTempGrid, tField, 0)
                    If tReturn.Length > 0 Then Return tReturn
                End If
            End If
            Return ""
        Catch ex As Exception
            MSG_ERROR("WhatIsClientVMBValue", ex.ToString)
        End Try
        Return ""
    End Function

    Public Function STR_convert_Macro(ByVal tLineIn As String) As String
        Try
            Dim tLineOut As String = tLineIn
            tLineOut = tLineOut.Replace("&ccyymmdd", rkutils.STR_format("TODAY", "ccyymmdd"))
            tLineOut = tLineOut.Replace("&ccyy", rkutils.STR_format("TODAY", "ccyy"))
            tLineOut = tLineOut.Replace("&cc", rkutils.STR_format("TODAY", "cc"))
            tLineOut = tLineOut.Replace("&yy", rkutils.STR_format("TODAY", "yy"))
            tLineOut = tLineOut.Replace("&mm", rkutils.STR_format("TODAY", "mm"))
            tLineOut = tLineOut.Replace("&dd", rkutils.STR_format("TODAY", "dd"))
            Return tLineOut
        Catch ex As Exception
            MSG_ERROR("STR_convert_Macro", ex.ToString)
            Return ""
        End Try
    End Function

    Public Function DataGridViewContains2(ByVal gGrid As DataGridView, ByVal tField As String, ByVal tValue As String, ByVal tField2 As String, ByVal tValue2 As String, ByVal bUpper As Boolean) As Integer
        Try
            Dim i1Grid As Integer
            For i1Grid = 0 To gGrid.Rows.Count - 1
                If bUpper Then
                    If rkutils.DataGridView_ValueByColumnName(gGrid, tField.ToUpper, i1Grid).Trim = tValue.ToUpper Then
                        If rkutils.DataGridView_ValueByColumnName(gGrid, tField2.ToUpper, i1Grid).Trim = tValue2.ToUpper Then
                            'sReturnValue = rkutils.DataGridView_ValueByColumnName(gGrid, sReturnField.ToUpper, i1Grid).Trim
                            Return i1Grid
                        End If
                    End If
                Else
                    If rkutils.DataGridView_ValueByColumnName(gGrid, tField, i1Grid).Trim = tValue Then
                        If rkutils.DataGridView_ValueByColumnName(gGrid, tField2, i1Grid).Trim = tValue2 Then
                            'sReturnValue = rkutils.DataGridView_ValueByColumnName(gGrid, sReturnField, i1Grid).Trim
                            Return i1Grid
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MSG_ERROR("DataGridViewContains2", ex.ToString)
        End Try
        Return -1
    End Function

    Public Sub TRACKS_update(ByVal SQLConnectionString As String, ByVal SQLuser As String, ByVal tCLIENT As String, ByVal tLOCX As String, ByVal tUNIQUE As String, ByVal tTYPE As String, ByVal tCOMMENT As String, ByVal TDate As String, ByVal TBy As String)
        Try
            '**********************************************************
            '* 2016-12-20 RFK:
            '* 2021-11-23 RFK: Cleaned Up Code
            Dim msSQLCommandString As String = ""
            If tUNIQUE.Length > 0 Then
                '
            Else
                msSQLCommandString = "INSERT INTO RevMD.dbo.tracks"
                msSQLCommandString += " (track_date"
                msSQLCommandString += ", track_by"
                msSQLCommandString += ", comment"
                msSQLCommandString += ", type"
                msSQLCommandString += ", client"
                msSQLCommandString += ", locx"
                msSQLCommandString += ") values("
                msSQLCommandString += "'" + TDate + "'"
                msSQLCommandString += ",'" + TBy + "'"
                msSQLCommandString += ",'" + tCOMMENT + "'"
                msSQLCommandString += ",'" + Left(tTYPE, 1) + "'"
                msSQLCommandString += ",'" + Left(tCLIENT, 20) + "'"
                msSQLCommandString += ",'" + Left(tLOCX, 20) + "'"
                msSQLCommandString += ")"
                DB_COMMAND("MSSQL", SQLConnectionString, SQLuser, msSQLCommandString)
            End If
        Catch ex As Exception
            'MsgBox("TRACKS_update", ex.ToString)
        End Try
    End Sub

    Public Function NOTES_ADD(ByVal tDB As String, ByVal SQLConnectionString As String, ByVal SQLuser As String, ByVal User400 As String, ByVal gGRID As DataGridView, ByVal tLOCX As String, ByVal tNUM As String, ByVal tMSGC As String, ByVal tSTAT As String, ByVal tContactCode As String, ByVal tBalance As String, ByVal tMatchBalance As String, ByVal tMESSAGE As String, ByVal TDate As String, ByVal TBy As String, ByVal tFree30 As String) As Boolean
        Try
            Select Case tDB
                Case "DB2"
#If DB2 = 1 Then
                    '******************************************************************************
                    '* 2016-12-20 RFK: Changed to RAPSTA
                    '* 2021-11-22 RFK: Cleaned Up Code
                    Dim iMSGnumber As Integer = rkutils.NOTES_MAXNUMBER(tDB, SQLConnectionString, SQLuser, User400, gGRID, tLOCX) + 1
                    Dim sRapSta As String = rkutils.SQL_READ_FIELD(gGRID, tDB, "RAPSTA", SQLConnectionString, SQLuser, "SELECT RARSTA,RAPSTA,RABALD FROM ROIDATA.RACCTP WHERE RALOCX='" + tLOCX + "'")
                    '******************************************************************************
                    If IsDate(TDate) = False Then TDate = Date.Now.ToString
                    If rkutils.STR_format(TDate, "ccyy") <= "1900" Then TDate = Date.Now.ToString
                    '******************************************************************************
                    Dim sSQL As String = "INSERT INTO ROIDATA.RGMSGP "
                    sSQL += " (RGLOCX, RGMSG#, RGMSGC, RGMON, RGDAY, RGYEAR, RGTIME, RGUSER"
                    sSQL += ", RGMSG, RGFR30, RGRNA, RGPRST, RGSTAT, RGCOCD, RGLNA, RGBAL"
                    sSQL += ") VALUES("
                    sSQL += "" + tLOCX                                                                                                          'RGLOCX
                    sSQL += "," + iMSGnumber.ToString.Trim                                                                                      'RGMSG#
                    sSQL += ",'" + rkutils.STR_TRIM(tMSGC, 2) + "'"                                                                             'RGMSGC
                    sSQL += "," + STR_format(TDate, "mm")                                                                                       'RGMON
                    sSQL += "," + STR_format(TDate, "dd")                                                                                       'RGDAY
                    sSQL += "," + STR_format(TDate, "ccyy")                                                                                     'RGYEAR
                    sSQL += "," + STR_format(TDate, "HHMM")                                                                                     'RGTIME
                    sSQL += ",'" + rkutils.STR_TRIM(User400, 6) + "'"                                                                           'RGUSER
                    sSQL += ",'" + rkutils.STR_TRIM(rkutils.STR_format(tMESSAGE, "VALID"), 87) + "'"                                            'RGMSG
                    sSQL += ",'" + rkutils.STR_TRIM(rkutils.STR_format(tFree30, "VALID"), 30) + "'"                                             'RGFR30
                    sSQL += "," + rkutils.STR_TRIM("0", 3)                                                                                      'RGRNA
                    sSQL += ",'" + rkutils.STR_TRIM(sRapSta, 3) + "'"                                                                           'RGPRST
                    sSQL += ",'" + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(gGRID, "RARSTA", 0), 3) + "'"                        'RGSTAT
                    sSQL += ",'" + rkutils.STR_TRIM(tContactCode, 1) + "'"                                                                      'RGCOCD
                    sSQL += "," + rkutils.STR_TRIM("0", 3)                                                                                      'RGLNA
                    sSQL += "," + rkutils.STR_TRIM(rkutils.STR_format(rkutils.DataGridView_ValueByColumnName(gGRID, "RABALD", 0), "0"), 12)     'RGBAL
                    sSQL += ")"
                    rkutils.DB_COMMAND(tDB, SQLConnectionString, SQLuser, sSQL)
                    rkutils.NOTES_PLUS(tDB, SQLConnectionString, SQLuser, User400, gGRID, tLOCX, "M")
                    '**********************************************************
                    Return True
#End If
                Case "MSSQL"
                    Dim msSQLCommandString As String = ""
            End Select
        Catch ex As Exception
            MSG_ERROR("NOTES_ADD", ex.ToString)
        End Try
        Return False
    End Function

    Public Function LOCX_STATUS(ByVal tDB As String, ByVal DB2SQLConnectionString As String, ByVal DB2SQLuser As String, ByVal MSSQLConnectionString As String, ByVal MSSQLuser As String, ByVal gGRID As DataGridView, ByVal gGRID2 As DataGridView, ByVal tLOCX As String, ByVal tContactCode As String, ByVal tRAC As String, ByVal tSTATUS As String, ByVal tNOTE As String, ByVal TDate As String, ByVal TBy As String, ByVal tModule As String) As Boolean
        Try
            Dim swOK As Boolean = False
            Select Case tDB
                Case "DB2"
#If DB2 = 1 Then
                    '**********************************************************
                    '* 2021-11-23 RFK: VERIFY VALID LOCX
                    If tLOCX <= 0 Then
                        eProcess.MsgStatus("***** ERROR ***** aoProcessor STATUS INVALID LOCX=" + tLOCX, False)
                        Return False
                    End If
                    '**********************************************************
                    '* 2015-09-29 RFK: corrected no need for SELECT *
                    Dim db2SQLCommandString As String = "SELECT RARSTA, RACLOS, RAMTTP, RAINB#, RAOUT#, RATOTC"
                    Dim sSQLcheck As String = ""
                    db2SQLCommandString += ", RATCAL, RAATMP, RACONT, RAADMI, RAATP2, RASTATS, RABALD, RAMBAL"
                    db2SQLCommandString += ", RASTDT, RASTTM"   '2021-03-10 RFK: For Older than current
                    db2SQLCommandString += " FROM ROIDATA.RACCTP WHERE RALOCX='" + tLOCX + "'"
                    'Process.MsgStatus(db2SQLCommandString, False)
                    '**********************************************************
                    Dim tCurrentStatus As String = rkutils.SQL_READ_FIELD(gGRID, tDB, "RARSTA", DB2SQLConnectionString, DB2SQLuser, db2SQLCommandString).Trim
                    '**********************************************************
                    '* RFK:  GridView1 now contains the required Fields so no need to SELECT again
                    Dim tMatchType As String = DataGridView_ValueByColumnName(gGRID, "RAMTTP", 0)
                    Dim tNumIn As String = DataGridView_ValueByColumnName(gGRID, "RAINB#", 0)
                    Dim tNumOut As String = DataGridView_ValueByColumnName(gGRID, "RAOUT#", 0)
                    Dim tNumCalls As String = DataGridView_ValueByColumnName(gGRID, "RATOTC", 0)
                    Dim tNumCallAttempts As String = DataGridView_ValueByColumnName(gGRID, "RATCAL", 0)
                    Dim tNumAttempts As String = DataGridView_ValueByColumnName(gGRID, "RAATMP", 0)
                    Dim tNumContacts As String = DataGridView_ValueByColumnName(gGRID, "RACONT", 0)
                    Dim tNumAdmin As String = DataGridView_ValueByColumnName(gGRID, "RAADMI", 0)
                    Dim tNumMessages As String = DataGridView_ValueByColumnName(gGRID, "RAATP2", 0)
                    Dim tRABALD As String = DataGridView_ValueByColumnName(gGRID, "RABALD", 0)
                    Dim tRAMBAL As String = DataGridView_ValueByColumnName(gGRID, "RAMBAL", 0)
                    Dim tRASTATS As String = DataGridView_ValueByColumnName(gGRID, "RASTATS", 0).Trim
                    Dim sRASTDT As String = DataGridView_ValueByColumnName(gGRID, "RASTDT", 0).Trim
                    Dim sRASTTM As String = DataGridView_ValueByColumnName(gGRID, "RASTTM", 0).Trim
                    '**********************************************************
                    '* 2016-12-21 RFK: Look up RARSTA status for STCALD
                    Select Case TBy
                        Case "Dialer"
                            '**************************************************
                            '* 2021-05-26 RFK: STCALD (INCORRECT - Match Away)
                            '* 2021-05-26 RFK: STDRDT (Dialer Status Over)
                            db2SQLCommandString = "SELECT S.STRECA,S.STDAYS,S.STCALD,S.STDRDT FROM ROIDATA.STATP S WHERE S.STSTAT='" + tCurrentStatus + "' AND STMTTP='" + tMatchType + "'"
                            Dim sSTDRDT As String = SQL_READ_FIELD(gGRID, "DB2", "STDRDT", DB2SQLConnectionString, DB2SQLuser, db2SQLCommandString).Trim
                            If sSTDRDT <> "Y" Then
                                eProcess.MsgStatus("Dialer can NOT status over [" + tCurrentStatus + "] STDRDT:" + DataGridView_ValueByColumnName(gGRID, "STDRDT", 0) + "]", True)
                                '**********************************************
                                '* 2019-07-10 RFK: If Already in Notes do NOT do
                                db2SQLCommandString = "SELECT * FROM ROIDATA.RGMSGP WHERE RGLOCX='" + tLOCX + "' AND RGMSG='" + "CAN NOT STATUS OVER " + tCurrentStatus + " with " + Trim(tSTATUS) + "'"
                                db2SQLCommandString += " AND RGYEAR='" + STR_format(TDate, "ccyy") + "' AND RGMON='" + STR_format(TDate, "mm") + "' AND RGDAY='" + STR_format(TDate, "dd") + "'"
                                SQL_READ_FIELD(gGRID2, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, db2SQLCommandString)
                                If gGRID2.Rows.Count <= 1 Then
                                    '**********************************************
                                    NOTES_ADD("DB2", DB2SQLConnectionString, DB2SQLuser, TBy, gGRID, tLOCX, "1", tMatchType, Trim(tSTATUS), tContactCode, tRABALD, tRAMBAL, "CAN NOT STATUS OVER " + tCurrentStatus + " with " + tSTATUS, TDate, TBy, "")
                                    '**********************************************
                                    '* 2017-03-01 RFK: UPDATE RATOTC / RAATMP (Attempts)
                                    db2SQLCommandString = "UPDATE ROIDATA.RACCTP"
                                    If Val(tNumCalls) + 1 > 99 Then
                                        db2SQLCommandString += " SET RATOTC=99"                                         'Total Number of Calls
                                    Else
                                        db2SQLCommandString += " SET RATOTC=" + Trim(Str(Val(tNumCalls) + 1))           'Total Number of Calls
                                    End If
                                    If Val(tNumCallAttempts) + 1 > 999 Then
                                        db2SQLCommandString += ", RATCAL=999"                                           'Total Call Attempts
                                    Else
                                        db2SQLCommandString += ", RATCAL=" + Trim(Str(Val(tNumCallAttempts) + 1))       'Total Call Attempts
                                    End If
                                    If Val(tNumAttempts) + 1 > 99 Then
                                        db2SQLCommandString += ", RAATMP=99"                                            'Attempts
                                    Else
                                        db2SQLCommandString += ", RAATMP=" + Str(Val(tNumAttempts) + 1)                 'Attempts
                                    End If
                                    If Val(tNumOut) + 1 > 999 Then
                                        db2SQLCommandString += ", RAOUT#=999"                                           'Outbound
                                    Else
                                        db2SQLCommandString += ", RAOUT#=" + Trim(Str(Val(tNumOut) + 1))                'Outbound
                                    End If
                                    db2SQLCommandString += rkutils.WhereAnd(db2SQLCommandString, "RALOCX=" + tLOCX)
                                    '**********************************************
                                    'eProcess.MsgStatus(db2SQLCommandString, False)
                                    DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, db2SQLCommandString)
                                    '**********************************************
                                Else
                                    eProcess.MsgStatus("ALREADY NOTED TODAY ... Dialer can NOT status over [" + tCurrentStatus + "] STCALD:" + DataGridView_ValueByColumnName(gGRID, "STCALD", 0) + "]", True)
                                End If
                                Return True
                            End If
                    End Select
                    '**********************************************************
                    '* 2012-02-01 RFK: check STRECA for Y to set RACLOS
                    '* 2012-04-05 RFK: checking STATUS
                    db2SQLCommandString = "SELECT S.STRECA,S.STDAYS,S.STCALD FROM ROIDATA.STATP S WHERE S.STSTAT='" + Trim(tSTATUS) + "' AND STMTTP='" + tMatchType + "'"
                    'eProcess.MsgStatus(db2SQLCommandString, False)
                    Dim tSTRECA As String = SQL_READ_FIELD(gGRID, "DB2", "STRECA", DB2SQLConnectionString, DB2SQLuser, db2SQLCommandString).Trim
                    '**********************************************************
                    '* 2018-01-25 RFK: Is it a valid status code
                    If gGRID.Rows.Count < 1 Then
                        eProcess.MsgStatus("UNABLE TO STATUS, INVALID STATUS=" + tSTATUS + "]STMTTP=" + tMatchType + "]", False)
                        Return False
                    End If
                    '**********************************************************
                    '* 2019-07-10 RFK: If Already in Notes do NOT do
                    db2SQLCommandString = "SELECT * FROM ROIDATA.RGMSGP WHERE RGLOCX='" + tLOCX + "' AND RGSTAT='" + Trim(tSTATUS) + "'"
                    db2SQLCommandString += " AND RGYEAR='" + STR_format(TDate, "ccyy") + "' AND RGMON='" + STR_format(TDate, "mm") + "' AND RGDAY='" + STR_format(TDate, "dd") + "'"
                    'eProcess.MsgStatus(db2SQLCommandString, False)
                    SQL_READ_FIELD(gGRID2, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, db2SQLCommandString)
                    If gGRID2.Rows.Count > 1 Then
                        Select Case TBy
                            Case "Dialer"
                                '**********************************************
                                '* 2021-11-01 RFK: DO NOT ALLOW FOR DIALER
                                eProcess.MsgStatus("UNABLE TO STATUS, ALREADY A STATUS TODAY=" + Trim(tSTATUS) + "]STMTTP=" + tMatchType + "]", True)
                                eProcess.MsgStatus(tLOCX + " " + gGRID2.Rows.Count.ToString, True)
                                TRACKS_update(MSSQLConnectionString, MSSQLuser, "", tLOCX, "", "F", "UNABLE TO STATUS, ALREADY A STATUS TODAY=" + Trim(tSTATUS) + "]STMTTP=" + tMatchType + "]", TDate, TBy)
                                Return False
                            Case Else
                                '**********************************************
                                '* 2021-11-01 RFK: [QUERIES AND COMMANDS SHOULD ALLOW]
                        End Select
                    End If
                    '**********************************************************
                    Dim tSTDAYS As String = DataGridView_ValueByColumnName(gGRID, "STDAYS", 0)
                    If Val(tSTDAYS) <= 0 Then tSTDAYS = "0"
                    '**********************************************************
                    '* 2012-04-05 RFK: checking STATUS
                    Dim tAttemptCTR As String = SQL_READ_FIELD(gGRID, "MSSQL", "CTR_ATTEMPT", MSSQLConnectionString, MSSQLuser, "SELECT S.CTR_ATTEMPT, S.CTR_CONTACT, S.CTR_ADMIN,S.CTR_MESSAGES,S.STATUS,S.TYPE FROM RevMD.dbo.STATUS S WHERE S.STATUS='" + Trim(tSTATUS) + "' AND S.TYPE='" + tMatchType + "'").Trim
                    '**********************************************************
                    '* RFK:  GridView1 now contains the required Fields so no need to SELECT again
                    Dim tContactCTR As String = DataGridView_ValueByColumnName(gGRID, "CTR_CONTACT", 0)
                    Dim tAdminCTR As String = DataGridView_ValueByColumnName(gGRID, "CTR_ADMIN", 0)
                    Dim tMessagesCTR As String = DataGridView_ValueByColumnName(gGRID, "CTR_MESSAGES", 0)
                    '**********************************************************
                    If IsDate(TDate) = False Or STR_format(TDate, "ccyymmdd") = "19000101" Then TDate = STR_format("TODAY", "mm/dd/ccyy HH:MM:SS")
                    '**********************************************************
                    db2SQLCommandString = "UPDATE ROIDATA.RACCTP"
                    '**********************************************************
                    db2SQLCommandString += " SET RACHGI='Y'"                                            'Change Indicator
                    sSQLcheck = "SELECT RALOCX"
                    sSQLcheck += " FROM ROIDATA.RACCTP WHERE RACHGI='Y'"                                'Change Indicator
                    '**********************************************************
                    '* 2021-03-10 RFK: Only if the Status is NEWER than last status date.
                    If STR_format(TDate, "ccyymmdd") > sRASTDT Then
                        swOK = True
                    Else
                        If STR_format(TDate, "ccyymmdd") = sRASTDT And STR_format(TDate, "HHMMSS") > STR_format(sRASTTM, "HHMMSS") Then
                            swOK = True
                        Else
                            swOK = False
                            eProcess.MsgStatus("swOK=" + swOK.ToString + " [" + STR_format(TDate, "ccyymmdd") + "][" + STR_format(TDate, "HHMMSS") + "][" + STR_format(sRASTTM, "HHMMSS") + "]", True)
                            TRACKS_update(MSSQLConnectionString, MSSQLuser, "", tLOCX, "", "F", "UNABLE TO UPDATE LAST STATUS INFORMATON A NEWER STATUS EXISTS", "", "")
                        End If
                    End If
                    '**********************************************************
                    '* 2021-03-11 RFK:
                    If swOK Then
                        db2SQLCommandString += ", RAPSTA='" + tCurrentStatus + "'"                                                      'Previous STATUS
                        sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RAPSTA='" + tCurrentStatus + "'")                                     'Previous STATUS
                        db2SQLCommandString += ", RARSTA='" + Trim(tSTATUS) + "'"                                                       'STATUS
                        sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RARSTA='" + Trim(tSTATUS) + "'")                                      'STATUS
                        db2SQLCommandString += ", RASTDT='" + STR_format(TDate, "ccyymmdd") + "'"                                       'Status Date
                        sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RASTDT='" + STR_format(TDate, "ccyymmdd") + "'")                      'Status Date
                        db2SQLCommandString += ", RASTTM='" + STR_format(TDate, "HH:MM:SS") + "'"                                       'Status Time
                        'sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RASTTM='" + STR_format(TDate, "HH:MM:SS") + "'")                      'Status Time   (Need to assign to variable 1st incase the MM/SS changes
                        db2SQLCommandString += ", RASNAM='" + STR_TRIM(TBy, 20) + "'"                                                   'Status By
                        sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RASNAM='" + STR_TRIM(TBy, 20) + "'")                                  'Status By
                        '******************************************************
                        '* 2012-02-01 RFK: Check STRECA for Y to set RACLOS
                        If tSTRECA = "Y" Then
                            db2SQLCommandString += ", RACLOS='C'"                                                                       'CLOSED 
                            sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RACLOS='C'")                                                      'CLOSED 
                        Else
                            db2SQLCommandString += ", RACLOS=''"                                                                        'CLOSED (NO)
                            sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RACLOS=''")                                                       'CLOSED (NO)
                        End If
                        '******************************************************
                        '* 2012-02-22 RFK: Set the NextData from NextDays
                        Dim tNextDate As String = ""
                        If Val(tSTDAYS) > 0 Then
                            tNextDate = Convert.ToString(DateAdd(DateInterval.Day, Val(tSTDAYS), Today))
                            If IsDate(tNextDate) Then
                                db2SQLCommandString += ", RARNMO=" + Trim(Str(Val(Convert.ToDateTime(tNextDate).Month.ToString)))                           'Review Next Month
                                sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RARNMO=" + Trim(Str(Val(Convert.ToDateTime(tNextDate).Month.ToString))))          'Review Next Month
                                db2SQLCommandString += ", RARNDY=" + Trim(Str(Val(Convert.ToDateTime(tNextDate).Day.ToString)))                             'Review Next Day
                                sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RARNDY=" + Trim(Str(Val(Convert.ToDateTime(tNextDate).Day.ToString))))            'Review Next Day
                                db2SQLCommandString += ", RARNYR=" + Trim(Str(Val(Convert.ToDateTime(tNextDate).Year.ToString)))                            'Review Next Year
                                sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RARNYR=" + Trim(Str(Val(Convert.ToDateTime(tNextDate).Year.ToString))))           'Review Next Year
                                db2SQLCommandString += " ,RANCDT='" + STR_format(tNextDate, "mmddccyy") + "'"                                               'Next Call Date
                                sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RANCDT='" + STR_format(tNextDate, "mmddccyy") + "'")                              'Next Call Date
                            End If
                        End If
                        '******************************************************
                    End If
                    '**********************************************************
                    db2SQLCommandString += ", RALCMO=" + Now.Month.ToString                                                                               'Change Date Month
                    sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RALCMO=" + Now.Month.ToString)                                                              'Change Date Month
                    db2SQLCommandString += ", RALCDY=" + Now.Day.ToString                                                                                 'Change Date DAy
                    sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RALCDY=" + Now.Day.ToString)                                                                'Change Date DAy
                    db2SQLCommandString += ", RALCYR=" + Now.Year.ToString                                                                                'Change Date Year
                    sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RALCYR=" + Now.Year.ToString)                                                               'Change Date Year
                    '**********************************************************
                    '* 2017-10-09 RFK: STATUS LIST (No Duplicates)
                    If tRASTATS.Contains(Trim(tSTATUS)) = False Then
                        db2SQLCommandString += ", RASTATS='" + STR_TRIM((tRASTATS + " " + tSTATUS), 50) + "'"
                        sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RASTATS='" + STR_TRIM((tRASTATS + " " + tSTATUS), 50) + "'")
                    End If
                    '**********************************************************
                    Select Case STR_TRIM(tContactCode, 1)
                        Case "I"
                            db2SQLCommandString += ", RAINB#=" + Trim(Str(Val(tNumIn) + 1))                                                                 'Inbound
                            sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RAINB#=" + Trim(Str(Val(tNumIn) + 1)))                                                'Inbound
                            db2SQLCommandString += ", RATOTC=" + Trim(Str(Val(tNumCalls) + 1))                                                              'Total Number of Calls
                            sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RATOTC=" + Trim(Str(Val(tNumCalls) + 1)))                                             'Total Number of Calls
                        Case "O"
                            db2SQLCommandString += ", RAOUT#=" + Trim(Str(Val(tNumOut) + 1))                                                                'Outbound
                            sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RAOUT#=" + Trim(Str(Val(tNumOut) + 1)))                                               'Outbound
                            db2SQLCommandString += ", RATOTC=" + Trim(Str(Val(tNumCalls) + 1))                                                              'Total Number of Calls
                            sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RATOTC=" + Trim(Str(Val(tNumCalls) + 1)))                                             'Total Number of Calls
                            db2SQLCommandString += ", RATCAL=" + Trim(Str(Val(tNumCallAttempts) + 1))                                                       'Total Call Attempts
                            sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RATCAL=" + Trim(Str(Val(tNumCallAttempts) + 1)))                                      'Total Call Attempts
                        Case "R"
                            '
                    End Select
                    '*****************************************************************************
                    '* 2012-04-05 RFK:
                    '* 2012-05-04 RFK: Messages Counter (Using Phone2 CTR)
                    If tAttemptCTR = "Y" Then
                        If Val(tNumAttempts) + 1 < 100 Then
                            db2SQLCommandString += ", RAATMP=" + Str(Val(tNumAttempts) + 1)
                            sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RAATMP=" + Str(Val(tNumAttempts) + 1))
                        Else
                            db2SQLCommandString += ", RAATMP=99"
                            sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RAATMP=99")
                        End If
                    End If
                    If tContactCTR = "Y" Then
                        If Val(tNumContacts) + 1 < 100 Then
                            db2SQLCommandString += ", RACONT=" + Str(Val(tNumContacts) + 1)
                            sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RACONT=" + Str(Val(tNumContacts) + 1))
                        Else
                            db2SQLCommandString += ", RACONT=99"
                            sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RACONT=99")
                        End If
                    End If
                    If tAdminCTR = "Y" Then
                        If Val(tNumAdmin) + 1 < 100 Then
                            db2SQLCommandString += ", RAADMI=" + Str(Val(tNumAdmin) + 1)
                            sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RAADMI=" + Str(Val(tNumAdmin) + 1))
                        Else
                            db2SQLCommandString += ", RAADMI=99"
                            sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RAADMI=99")
                        End If
                    End If
                    If tMessagesCTR = "Y" Then
                        If Val(tNumMessages) + 1 < 100 Then
                            db2SQLCommandString += ", RAATP2=" + Str(Val(tNumMessages) + 1)                                         'Number of Messages (Using Phone2 CTR)
                            sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RAATP2=" + Str(Val(tNumMessages) + 1))                        'Number of Messages (Using Phone2 CTR)
                        Else
                            db2SQLCommandString += ", RAATP2=99"
                            sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RAATP2=99")
                        End If
                    End If
                    '**********************************************************
                    db2SQLCommandString += " WHERE RALOCX=" + tLOCX
                    sSQLcheck += rkutils.WhereAnd(sSQLcheck, "RALOCX=" + tLOCX)
                    '**********************************************************
                    'eProcess.MsgStatus(db2SQLCommandString, False)
                    DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, db2SQLCommandString)
                    '**********************************************************
                    '* 2012-04-05 RFK: changed the MC (MessageCode to AI / AO / AR
                    Dim tMessageCode As String = "A" + Left(tContactCode, 1)
                    '**********************************************************
                    '* 2021-11-18 RFK: VERIFY THAT THE UPDATE/SET OCCURRED 
                    If SQL_READ_DATAGRID(gGRID, "DB2", "RALOCX", DB2SQLConnectionString, DB2SQLuser, sSQLcheck) Then
                        '******************************************************
                        '* Good To Go!
                        eProcess.MsgStatus("" + db2SQLCommandString, False)
                    Else
                        eProcess.MsgStatus("***** ERROR ***** aoProcessor STATUS FAILURE LOCX=" + tLOCX, False)
                        eProcess.MsgStatus("***** ERROR ***** " + db2SQLCommandString, False)
                        eProcess.MsgStatus("***** ERROR ***** " + sSQLcheck, False)
                        rkutils.EMAILIT(MSSQLConnectionString, MSSQLuser, "DoNotReply@AnnuityHealth.com", "aoProcessor", "Ryan@AnnuityHealth.com", "iIncident", "", "aoProcessor STATUS FAILURE LOCX=" + tLOCX, sSQLcheck, "", "")
                    End If
                    '**********************************************************
                    If tNOTE.Trim.Length > 0 Then
                        NOTES_ADD("DB2", DB2SQLConnectionString, DB2SQLuser, TBy, gGRID, tLOCX, "1", tMessageCode, Trim(tSTATUS), tContactCode, tRABALD, tRAMBAL, "STATUSED: " + tNOTE, TDate, TBy, "")
                    Else
                        '******************************************************
                        '* 2021-11-23 RFK: STMTTP (WAS A) / corrected
                        Dim tStatusDescr As String = SQL_READ_FIELD(gGRID, "DB2", "STDESC", DB2SQLConnectionString, DB2SQLuser, "SELECT STDESC FROM ROIDATA.STATP WHERE STMTTP='" + tMatchType + "' AND STSTAT='" + Trim(tSTATUS) + "'")
                        NOTES_ADD("DB2", DB2SQLConnectionString, DB2SQLuser, TBy, gGRID, tLOCX, "1", tMessageCode, Trim(tSTATUS), tContactCode, tRABALD, tRAMBAL, "STATUSED: " + tStatusDescr, TDate, TBy, "")
                    End If
                    '**********************************************************
                    Return True
#End If
            End Select
        Catch ex As Exception
            MSG_ERROR("LOCX_STATUS", ex.ToString)
        End Try
        Return False
    End Function

    Public Function Encrypt(sVersion As String, clearText As String, EncryptionKey As String) As String
        Try
            Select Case sVersion
                Case "VB6"
                    '*****************************
                    '* 2015-10-19 RFK: vb6 version
                    Dim iLen As Integer, iX As Integer, iX2 As Integer
                    iLen = Len(EncryptionKey)
                    iX2 = 1
                    Dim sSTR As String
                    sSTR = ""
                    For iX = 1 To Len(clearText)
                        sSTR = sSTR + Trim(Str(Asc(Mid(clearText, iX, 1)) + Asc(Mid(EncryptionKey, iX2, 1)))).PadLeft(3, "0")
                        iX2 = iX2 + 1
                        If iX2 > iLen Then iX2 = 1
                    Next
                    Return sSTR
                Case Else
                    '*****************************
                    '* RFK: VB.NET VERSION (much more secure)
                    Dim clearBytes As Byte() = Encoding.Unicode.GetBytes(clearText)
                    Using encryptor As Aes = Aes.Create()
                        Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, &H65, &H64, &H76, &H65, &H64, &H65, &H76})
                        encryptor.Key = pdb.GetBytes(32)
                        encryptor.IV = pdb.GetBytes(16)
                        Using ms As New MemoryStream()
                            Using cs As New CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
                                cs.Write(clearBytes, 0, clearBytes.Length)
                                cs.Close()
                            End Using
                            clearText = Convert.ToBase64String(ms.ToArray())
                        End Using
                    End Using
            End Select
        Catch ex As Exception
            MSG_ERROR("Encrypt", ex.ToString)
        End Try
        Return clearText
    End Function

    Public Function Decrypt(sVersion As String, cipherText As String, EncryptionKey As String) As String
        Try
            Select Case sVersion
                Case "VB6"
                    '*****************************
                    '* 2015-10-19 RFK: vb6 version
                    Dim iLen As Integer, iX As Integer, iX2 As Integer
                    iLen = Len(EncryptionKey)
                    iX2 = 1
                    Dim sSTR As String
                    sSTR = ""
                    For iX = 1 To Len(cipherText) Step 3
                        sSTR = sSTR + Trim(Chr(Val(Mid(cipherText, iX, 3)) - Asc(Mid(EncryptionKey, iX2, 1))))
                        iX2 = iX2 + 1
                        If iX2 > iLen Then iX2 = 1
                    Next
                    Return sSTR
                Case Else
                    '*****************************
                    '* RFK: VB.NET VERSION (much more secure)
                    Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)
                    Using encryptor As Aes = Aes.Create()
                        Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D,
                         &H65, &H64, &H76, &H65, &H64, &H65,
                         &H76})
                        encryptor.Key = pdb.GetBytes(32)
                        encryptor.IV = pdb.GetBytes(16)
                        Using ms As New MemoryStream()
                            Using cs As New CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write)
                                cs.Write(cipherBytes, 0, cipherBytes.Length)
                                cs.Close()
                            End Using
                            cipherText = Encoding.Unicode.GetString(ms.ToArray())
                        End Using
                    End Using
            End Select
        Catch ex As Exception
            MSG_ERROR("Decrypt", ex.ToString)
        End Try
        Return cipherText
    End Function

End Module
