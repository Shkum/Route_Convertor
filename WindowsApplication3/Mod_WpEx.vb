Module Module2
    Public Function WpExt(fileType As String) As String
        On Error GoTo 1
        Select Case FileType


            Case "BVS"
                Dim c_WpName, c_Lat, c_Lon As String
                Dim cWP1, cWP, cWPtmp, cWPlast, cLine1, cLinelast As String
                Dim c_Str As Array, c_ChrPos As Integer
                Dim c_ML, c_Mline As String

                cLine1 = My.Resources.Res1.String4
                cWP1 = My.Resources.Res1.String5
                cWPtmp = My.Resources.Res1.String6
                cWPlast = My.Resources.Res1.String7
                cLinelast = My.Resources.Res1.String8
                
                For c_Count = 0 To Form1.ListBox1.Items.Count - 1
                    c_Str = Split(Form1.ListBox1.Items.Item(c_Count), vbTab)
                    c_WpName = c_Str(1)
                    c_Lat = c_Str(2)
                    c_ChrPos = InStr(c_Lat, Chr(176))
                    c_Lat = Left(c_Lat, c_ChrPos - 1) + (Val(Mid(c_Lat, c_ChrPos + 1, c_Lat.Length - c_ChrPos - 1)) / 60) & Right(c_Lat, 1)
                    If Right(c_Lat, 1) = "S" Then
                        c_Lat = "-" & Left(c_Lat, c_Lat.Length - 1)
                    Else
                        c_Lat = Left(c_Lat, c_Lat.Length - 1)
                    End If
                    c_Lat = Format(CDbl(c_Lat), "#.0000000000")
                    c_Lat = Replace(c_Lat, ",", ".")

                    c_Lon = c_Str(3)
                    c_ChrPos = InStr(c_Lon, Chr(176))
                    c_Lon = Left(c_Lon, c_ChrPos - 1) + (Val(Mid(c_Lon, c_ChrPos + 1, c_Lon.Length - c_ChrPos - 1)) / 60) & Right(c_Lon, 1)
                    If Right(c_Lon, 1) = "W" Then
                        c_Lon = "-" & Left(c_Lon, c_Lon.Length - 1)
                    Else
                        c_Lon = Left(c_Lon, c_Lon.Length - 1)
                    End If

                    c_Lon = Format(CDbl(c_Lon), "#.0000000000")
                    c_Lon = Replace(c_Lon, ",", ".")
                    
                    If c_Count=0 Then 
                        cWP1 = Replace(cWP1, "@@@@@", c_WpName)
                        cWP1 = Replace(cWP1, "^^^^^", c_Lat)
                        cWP1 = Replace(cWP1, "#####", c_Lon)

                    ElseIf c_Count = Form1.ListBox1.Items.Count - 1 Then
                        cWPlast = Replace(cWPlast, "@@@@@", c_WpName)
                        cWPlast = Replace(cWPlast, "^^^^^", c_Lat)
                        cWPlast = Replace(cWPlast, "#####", c_Lon)
                        
                    Else
                        cWP = Replace(cWPtmp, "^^^^^", c_Lat)
                        cWP = Replace(cWP, "#####", c_Lon)
                        c_ML = c_ML & cWP & vbCrLf
                    End If
                    
                Next

                c_Mline = cLine1 & vbCrLf & cWP1 & vbCrLf & c_ML & cWPlast & cLinelast
                Return c_Mline

                




            Case "RTE"

                Dim c_WpName, c_Lat, c_Lon, c_RlGc As String
                Dim c_line1, c_line2, c_mainLine, c_ML, c_Mline As String, c_ChrPos As Integer
                Dim c_Count As Integer, c_Str As Array
                c_line1 = My.Resources.Res1.String1
                c_mainLine = My.Resources.Res1.String2
                c_line2 = My.Resources.Res1.String3

                For c_Count = 0 To Form1.ListBox1.Items.Count - 1
                    c_Str = Split(Form1.ListBox1.Items.Item(c_Count), vbTab)
                    c_WpName = c_Str(1)
                    c_Lat = c_Str(2)
                    c_ChrPos = InStr(c_Lat, Chr(176))
                    c_Lat = Left(c_Lat, c_ChrPos - 1) + (Val(Mid(c_Lat, c_ChrPos + 1, c_Lat.Length - c_ChrPos - 1)) / 60) & Right(c_Lat, 1)
                    If Right(c_Lat, 1) = "S" Then
                        c_Lat = "-" & Left(c_Lat, c_Lat.Length - 1)
                    Else
                        c_Lat = Left(c_Lat, c_Lat.Length - 1)
                    End If
                    c_Lat = Replace(c_Lat, ",", ".")

                    c_Lon = c_Str(3)
                    c_ChrPos = InStr(c_Lon, Chr(176))
                    c_Lon = Left(c_Lon, c_ChrPos - 1) + (Val(Mid(c_Lon, c_ChrPos + 1, c_Lon.Length - c_ChrPos - 1)) / 60) & Right(c_Lon, 1)
                    If Right(c_Lon, 1) = "W" Then
                        c_Lon = "-" & Left(c_Lon, c_Lon.Length - 1)
                    Else
                        c_Lon = Left(c_Lon, c_Lon.Length - 1)
                    End If
                    c_Lon = Replace(c_Lon, ",", ".")


                    If c_Str(4) = "RL" Then
                        c_RlGc = "rhumbline"
                    Else
                        c_RlGc = "greatcircle"
                    End If

                    c_ML = Replace(c_mainLine, "@@@@@", c_WpName)
                    c_ML = Replace(c_ML, "%%%%%", c_RlGc)
                    c_ML = Replace(c_ML, "#####", c_Lat)
                    c_ML = Replace(c_ML, "$$$$$", c_Lon)
                    c_Mline = c_Mline & c_ML & vbCrLf
                Next

                c_Mline = c_line1 & vbCrLf & c_Mline & c_line2

                Return c_Mline





            Case "RT3"
                Dim c_WpName, c_Lat, c_Lon, c_RlGc As String
                Dim c_line1, c_line2, c_LINE3, c_line4, c_mainLine, c_MLine, c_ML, c_line3_1 As String

                Dim c_Count As Integer, c_Str As Array
                c_line1 = "<TSH_Route RtVersion=""3"" RtName=""?????"">TSH RtServer route data file. Info: amo." & vbCrLf & "<WayPoints WPCount=""+++++"">"
                c_line2 = "</WayPoints>" & vbCrLf & "<Calculations CalcCount=""1"">" & vbCrLf & "<Calculation CalcName=""BaseCalc"" CalcOptions=""0"" CalcForecast=""0"" CalcDone=""0"">""" & vbCrLf & "<WayPointExs>"
                c_LINE3 = "<WayPointEx ChangedData=""0"" TimeZone=""0"" ETA=""0"" ETD=""0"" Stay=""0"" TTG=""0"" TotalTime=""0"" Speed=""0""/>"
                c_line4 = "</WayPointExs>" & vbCrLf & "</Calculation>" & vbCrLf & "</Calculations>" & vbCrLf & "</TSH_Route>"
                c_mainLine = "<WayPoint WPName=""$$$$$"" LegType=""^^^^^"" RudderAngle=""0"" Lat=""@@@@@"" Lon=""%%%%%"" PortXTE=""0.100000001490116"" StbXTE=""0.100000001490116"" TurnRate=""0"" TurnRadius=""0.300000011920929"" ArrivalC=""0""/>"
                c_line1 = Replace(c_line1, "+++++", Form1.ListBox1.Items.Count)
                For c_Count = 0 To Form1.ListBox1.Items.Count - 1
                    c_Str = Split(Form1.ListBox1.Items.Item(c_Count), vbTab)
                    c_WpName = c_Str(1)
                    c_Lat = c_Str(2)
                    c_Lat = Val(Mid(c_Lat, 1, InStr(c_Lat, Chr(176)) - 1)) * 60 + Val(Mid(c_Lat, InStr(c_Lat, Chr(176)) + 1, c_Lat.Length - InStr(c_Lat, Chr(176)) - 1)) & Mid(c_Lat, c_Lat.Length, 1)
                    If Mid(c_Lat, c_Lat.Length, 1) = "S" Then
                        c_Lat = "-" & Mid(c_Lat, 1, c_Lat.Length - 1)
                    Else
                        c_Lat = Mid(c_Lat, 1, c_Lat.Length - 1)
                    End If
                    c_Lat = Replace(c_Lat, ",", ".")
                    c_Lon = c_Str(3)
                    c_Lon = Val(Mid(c_Lon, 1, InStr(c_Lon, Chr(176)) - 1)) * 60 + Val(Mid(c_Lon, InStr(c_Lon, Chr(176)) + 1, c_Lon.Length - InStr(c_Lon, Chr(176)) - 1)) & Mid(c_Lon, c_Lon.Length, 1)
                    If Mid(c_Lon, c_Lon.Length, 1) = "W" Then
                        c_Lon = "-" & Mid(c_Lon, 1, c_Lon.Length - 1)
                    Else
                        c_Lon = Mid(c_Lon, 1, c_Lon.Length - 1)
                    End If
                    c_Lon = Replace(c_Lon, ",", ".")
                    If c_Str(4) = "RL" Then
                        c_RlGc = "0"
                    Else
                        c_RlGc = "1"
                    End If
                    c_ML = Replace(c_mainLine, "$$$$$", c_WpName)
                    c_ML = Replace(c_ML, "^^^^^", c_RlGc)
                    c_ML = Replace(c_ML, "@@@@@", c_Lat)
                    c_ML = Replace(c_ML, "%%%%%", c_Lon)
                    c_MLine = c_MLine & c_ML & vbCrLf
                    c_line3_1 = c_line3_1 & vbCrLf & c_LINE3
                Next
                c_MLine = c_line1 & vbCrLf & c_MLine & c_line2 & c_line3_1 & vbCrLf & c_line4
                Return c_MLine


        End Select
1:
        Select Case Err.Number
            Case 0
            Case Else
                MsgBox("File not saved." & vbCrLf & "Error " & Err.Number & " - " & Err.Description, vbExclamation, "Error saving file")
        End Select
    End Function
End Module
