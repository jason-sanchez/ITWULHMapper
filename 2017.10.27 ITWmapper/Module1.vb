'20130401 - STAR_ITW Mapper for STAR Feeds
'20130408 - added code for multiple dg1 segments
'20130428 - Mods for STAR production feed.
'20140202 - mods for wave3 testing on cscsysfeed5
'20140205 - mod to use log file for errors instead of event viewer.
'20140215 - mod for multiple AL1 segments.
'20140325 - Don't create NVP file if MSH segment missing.
'20140817 - mods for W3 Production

'20150413 - VS2013 version

Imports System
Imports System.IO
Imports System.Collections
Imports System.data.SqlClient
Module Module1
    Private fullinipath As String = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory, "..\..\..\Configs\ULH\HL7Mapper.ini")) ' New test
    Public objIniFile As New INIFile(fullinipath) '20140817 - New Test
    'Public objIniFile As New INIFile("d:\W3Production\HL7Mapper.ini") '20140817 - Prod
    'Public objIniFile As New INIFile("C:\W3Feeds\HL7Mapper.ini") '20140817 - Test
    'Public objIniFile As New INIFile("C:\KY1 Test Environment\HL7Mapper.ini") '20140817 - Local
    Dim strInputDirectory As String = ""
    Dim strOutputDirectory As String = ""
    Dim strOutputSubDirectory As String = ""
    Dim strMapperFile As String = ""
    Dim strLogDirectory As String = ""
    Dim myHT As New Hashtable
    Dim filecounter As Integer = 0
    Public theFile As FileInfo  ' make sure this declaration outside try finally to be global to catch exception section.
    '20091119 - make thefile public for exception handling
    '20140325 - make sure record has an MSH segment
    Dim boolMSHExists As Boolean = False
    Sub Main()
        Try
            Dim strLTWOutput As String = ""
            Dim boolUseIt As Boolean = False '20081120
            'declarations for split function

            Dim dictNVP As New Hashtable
            Dim s As String
            Dim s1 As String
            'Dim connectionString As String = "server=(local);database=mapper;trusted_connection=true"
            'Dim connectionString As String = "Data Source=INSPIRON\SQLEXPRESS;Initial Catalog=mapper;Integrated Security=True"
            'Dim sql As String = ""
            'Dim dataReader As SqlDataReader
            'Dim myConnection As New SqlConnection(connectionString)
            'Dim objCommand As New SqlCommand
            'objCommand.Connection = myConnection
            'Dim myfile As StreamReader
            Dim dir As String

            Dim direct As String = objIniFile.GetString("Settings", "directory", "(none)") & ":\"
            Dim parent As String = objIniFile.GetString("Settings", "parentDir", "(none)") & "\"

            strInputDirectory = direct & parent & objIniFile.GetString("ITW", "ITWinputDir", "(none)") 'd:\feeds\HL7\itworks\
            strOutputDirectory = direct & parent & objIniFile.GetString("ITW", "ITWoutputdirectory", "(none)") 'd:\feeds\NVP\itworks\
            strMapperFile = direct & parent & objIniFile.GetString("ITW", "ITWmapper", "(none)")
            '20140205 - add logfile location
            strLogDirectory = direct & parent & objIniFile.GetString("Settings", "logs", "(none)")


            Dim delimStr As String = "|"
            Dim delimiter As Char() = delimStr.ToCharArray()
            Dim theFile As FileInfo
            Dim OBXCounter As Integer = 0
            Dim NTECounter As Integer = 0
            Dim IN1Counter As Integer = 0
            Dim IN2Counter As Integer = 0
            Dim IN3Counter As Integer = 0
            Dim ROLCounter As Integer = 0
            Dim ZCDCounter As Integer = 0
            Dim PIDCounter As Integer = 0
            Dim PV1Counter As Integer = 0
            '20100728
            Dim ZMICounter As Integer = 0
            '20130408 - added DG1 multiples for STAR
            Dim DG1Counter As Integer = 0
            '20140215
            Dim AL1Counter As Integer = 0
            '20170615 - for additional authcodes
            Dim ZGICounter As Integer = 0

            'declarations for stream reader
            Dim strLine As String = ""
            'setup directory
            Dim dirs As String() = Directory.GetFiles(strInputDirectory, "HL7.*")

            '20080714 create the reference hash table with mapper data
            CreateHashTable()

            For Each dir In dirs
                boolMSHExists = False '20140325
                filecounter = filecounter + 1
                If filecounter >= 201 Then Exit For
                OBXCounter = 0
                NTECounter = 0
                IN1Counter = 0
                IN2Counter = 0
                IN3Counter = 0
                ROLCounter = 0
                ZCDCounter = 0
                PIDCounter = 0
                PV1Counter = 0
                '20100728
                ZMICounter = 0
                '20130408
                DG1Counter = 0
                '20140215
                AL1Counter = 0
                '20170615
                ZGICounter = 0

                strLTWOutput = ""
                theFile = New FileInfo(dir)
                'LogFile.WriteLine(theFile.FullName)
                'If theFile.Extension <> ".$#$" Then

                '1.set up the streamreader to get a file
                'myfile = File.OpenText(dir)
                Dim myfile As StreamReader = New StreamReader(theFile.FullName)
                'LogFile.WriteLine(myfile)
                'and read the first line
                'strLine = myfile.ReadLine()

                Do
                    Dim myArray As String() = Nothing
                    Dim TestPos As Integer = 0
                    strLine = myfile.ReadLine()
                    Dim segId As String = ""
                    Dim segId2 As String = ""
                    Dim segIDFull As String = ""
                    Dim counter As Integer = 0

                    Dim segname As String = ""
                    'get the segment Id which is the first three Characters of the string
                    segId = Mid(strLine, 1, 3)
                    '20171018 added to handle instances where a ZGI segment is skipped
                    segId2 = Mid(strLine, 5, 1)

                    If segId = "MSH" Then
                        boolMSHExists = True '20140325
                        counter = +1
                        If counter = 1 Then
                            counter = +1
                        End If
                    End If

                    If segId = "OBX" Then
                        OBXCounter = OBXCounter + 1
                        If OBXCounter = 1 Then
                            OBXCounter = +1
                        End If
                    End If

                    If segId = "NTE" Then
                        NTECounter = NTECounter + 1
                        If NTECounter = 1 Then
                            NTECounter = +1
                        End If
                    End If

                    If segId = "IN1" Then
                        IN1Counter = IN1Counter + 1
                        If IN1Counter = 1 Then
                            IN1Counter = +1
                        End If
                    End If
                    If segId = "IN2" Then
                        IN2Counter = IN2Counter + 1
                        If IN2Counter = 1 Then
                            IN2Counter = +1
                        End If
                    End If
                    If segId = "IN3" Then
                        IN3Counter = IN3Counter + 1
                        If IN3Counter = 1 Then
                            IN3Counter = +1
                        End If
                    End If

                    If segId = "ROL" Then
                        ROLCounter = ROLCounter + 1
                        If ROLCounter = 1 Then
                            ROLCounter = +1
                        End If
                    End If
                    If segId = "ZCD" Then
                        ZCDCounter = ZCDCounter + 1
                        If ZCDCounter = 1 Then
                            ZCDCounter = +1
                        End If
                    End If
                    If segId = "PID" Then
                        PIDCounter = PIDCounter + 1
                        If PIDCounter = 1 Then
                            PIDCounter = +1
                        End If
                    End If
                    If segId = "PV1" Then
                        PV1Counter = PV1Counter + 1
                        If PV1Counter = 1 Then
                            PV1Counter = +1
                        End If
                    End If
                    '20100728
                    If segId = "ZMI" Then
                        ZMICounter = ZMICounter + 1
                        If ZMICounter = 1 Then
                            ZMICounter = +1
                        End If
                    End If
                    '20130408
                    If segId = "DG1" Then
                        DG1Counter = DG1Counter + 1
                        If DG1Counter = 1 Then
                            DG1Counter = +1
                        End If
                    End If

                    '20140215
                    If segId = "AL1" Then
                        AL1Counter = AL1Counter + 1
                        If AL1Counter = 1 Then
                            AL1Counter = +1
                        End If
                    End If
                    '20170605
                    If segId = "ZGI" Then
                        ZGICounter = ZGICounter + 1
                        If ZGICounter = 1 Then
                            ZGICounter = +1
                        End If
                        '20171018 added to handle instances where a ZGI segment is skipped
                        If CInt(segId2) > ZGICounter Then
                            ZGICounter = CInt(segId2)
                        End If
                    End If

                    'LogFile.WriteLine("---------------------------------------")
                    If strLine <> "" Then
                        myArray = strLine.Split(delimiter)
                        'add array key and item to hashtable
                        For Each s In myArray
                            'counter += 1

                            'If s <> "" Then
                            Dim mySubArray As String() = Nothing
                            mySubArray = s.Split("^")
                            segIDFull = segId & "_" & counter

                            If myHT.Item(segIDFull) <> "" Then
                                segname = myHT.Item(segIDFull)
                                boolUseIt = True ''20081120
                            Else
                                segname = segIDFull
                                boolUseIt = False ''20081120
                            End If
                            If boolUseIt = True Then '20081120
                                '=================================================================================
                                If segId = "OBX" Then
                                    'LogFile.Write(segname & "_" & OBXCounter & "=")
                                    strLTWOutput = strLTWOutput & segname & BuildSegCounter(OBXCounter) & "="


                                ElseIf segId = "NTE" Then
                                    'LogFile.Write(segname & "_" & OBXCounter & "=")
                                    strLTWOutput = strLTWOutput & segname & BuildSegCounter(NTECounter) & "="


                                ElseIf segId = "IN1" Then
                                    'LogFile.Write(segname & "_" & OBXCounter & "=")
                                    strLTWOutput = strLTWOutput & segname & BuildSegCounter(IN1Counter) & "="


                                ElseIf segId = "IN2" Then
                                    'LogFile.Write(segname & "_" & OBXCounter & "=")
                                    strLTWOutput = strLTWOutput & segname & BuildSegCounter(IN2Counter) & "="


                                ElseIf segId = "IN3" Then
                                    'LogFile.Write(segname & "_" & OBXCounter & "=")
                                    strLTWOutput = strLTWOutput & segname & BuildSegCounter(IN3Counter) & "="


                                ElseIf segId = "ROL" Then
                                    'LogFile.Write(segname & "_" & OBXCounter & "=")
                                    strLTWOutput = strLTWOutput & segname & BuildSegCounter(ROLCounter) & "="

                                ElseIf segId = "ZCD" Then
                                    'LogFile.Write(segname & "_" & OBXCounter & "=")
                                    strLTWOutput = strLTWOutput & segname & BuildSegCounter(ZCDCounter) & "="
                                ElseIf segId = "PID" Then
                                    'LogFile.Write(segname & "_" & OBXCounter & "=")
                                    strLTWOutput = strLTWOutput & segname & BuildSegCounter(PIDCounter) & "="
                                ElseIf segId = "PV1" Then
                                    'LogFile.Write(segname & "_" & OBXCounter & "=")
                                    strLTWOutput = strLTWOutput & segname & BuildSegCounter(PV1Counter) & "="
                                    '20100728
                                ElseIf segId = "ZMI" Then
                                    'LogFile.Write(segname & "_" & OBXCounter & "=")
                                    strLTWOutput = strLTWOutput & segname & BuildSegCounter(ZMICounter) & "="
                                    '20130408 - DG1 building for STAR
                                ElseIf segId = "DG1" Then

                                    strLTWOutput = strLTWOutput & segname & BuildSegCounter(DG1Counter) & "="

                                    '20140215
                                ElseIf segId = "AL1" Then

                                    strLTWOutput = strLTWOutput & segname & BuildSegCounter(AL1Counter) & "="
                                    '20170615
                                ElseIf segId = "ZGI" Then
                                    strLTWOutput = strLTWOutput & segname & BuildSegCounter(ZGICounter) & "="

                                Else
                                    'LogFile.Write(segname & "=")
                                    strLTWOutput = strLTWOutput & segname & "="
                                End If
                                '=================================================================================
                            End If '20081120

                            '20080721===========================================================================
                            If segIDFull = "MSH_5" Then
                                'strOutputSubDirectory = ""
                                'CreateOutputSubDirectory(s)


                            End If
                            '20080721 - end====================================================================
                            If boolUseIt Then '20081120
                                strLTWOutput = strLTWOutput & s & vbCrLf
                            End If '20081120
                            TestPos = InStr(1, s, "^")
                            If TestPos >= 0 Then
                                Dim subCounter As Integer = 0
                                For Each s1 In mySubArray
                                    subCounter += 1
                                    segIDFull = segId & "_" & counter & "_" & subCounter
                                    If myHT.Item(segIDFull) <> "" Then
                                        segname = myHT.Item(segIDFull)
                                        boolUseIt = True '20081120
                                    Else
                                        segname = segIDFull
                                        boolUseIt = False '20081120
                                    End If
                                    If boolUseIt Then '20081120
                                        '=================================================================================

                                        If segId = "OBX" Then
                                            'LogFile.Write(segname & "_" & OBXCounter & "=")
                                            strLTWOutput = strLTWOutput & segname & BuildSegCounter(OBXCounter) & "="


                                        ElseIf segId = "NTE" Then
                                            'LogFile.Write(segname & "_" & OBXCounter & "=")
                                            strLTWOutput = strLTWOutput & segname & BuildSegCounter(NTECounter) & "="


                                        ElseIf segId = "IN1" Then
                                            'LogFile.Write(segname & "_" & OBXCounter & "=")
                                            strLTWOutput = strLTWOutput & segname & BuildSegCounter(IN1Counter) & "="


                                        ElseIf segId = "IN2" Then
                                            'LogFile.Write(segname & "_" & OBXCounter & "=")
                                            strLTWOutput = strLTWOutput & segname & BuildSegCounter(IN2Counter) & "="


                                        ElseIf segId = "IN3" Then
                                            'LogFile.Write(segname & "_" & OBXCounter & "=")
                                            strLTWOutput = strLTWOutput & segname & BuildSegCounter(IN3Counter) & "="


                                        ElseIf segId = "ROL" Then
                                            'LogFile.Write(segname & "_" & OBXCounter & "=")
                                            strLTWOutput = strLTWOutput & segname & BuildSegCounter(ROLCounter) & "="

                                        ElseIf segId = "ZCD" Then
                                            'LogFile.Write(segname & "_" & OBXCounter & "=")
                                            strLTWOutput = strLTWOutput & segname & BuildSegCounter(ZCDCounter) & "="
                                        ElseIf segId = "PID" Then
                                            'LogFile.Write(segname & "_" & OBXCounter & "=")
                                            strLTWOutput = strLTWOutput & segname & BuildSegCounter(PIDCounter) & "="
                                        ElseIf segId = "PV1" Then
                                            'LogFile.Write(segname & "_" & OBXCounter & "=")
                                            strLTWOutput = strLTWOutput & segname & BuildSegCounter(PV1Counter) & "="

                                            '20100728
                                        ElseIf segId = "ZMI" Then
                                            'LogFile.Write(segname & "_" & OBXCounter & "=")
                                            strLTWOutput = strLTWOutput & segname & BuildSegCounter(ZMICounter) & "="

                                            '20130408 - STAR
                                        ElseIf segId = "DG1" Then
                                            'LogFile.Write(segname & "_" & OBXCounter & "=")
                                            strLTWOutput = strLTWOutput & segname & BuildSegCounter(DG1Counter) & "="

                                            '20140215
                                        ElseIf segId = "AL1" Then
                                            'LogFile.Write(segname & "_" & OBXCounter & "=")
                                            strLTWOutput = strLTWOutput & segname & BuildSegCounter(AL1Counter) & "="

                                            '20170615
                                        ElseIf segId = "ZGI" Then
                                            strLTWOutput = strLTWOutput & segname & BuildSegCounter(ZGICounter) & "="


                                        Else
                                            If boolUseIt Then '20081120
                                                strLTWOutput = strLTWOutput & segname & "="
                                            End If '20081120
                                        End If

                                        'LogFile.WriteLine(s1)
                                        strLTWOutput = strLTWOutput & s1 & vbCrLf
                                    End If '20081120
                                    '=================================================================================
                                Next
                            End If
                            'End If
                            counter += 1
                        Next

                    End If
                Loop Until (strLine Is Nothing)

                myfile.Close()

                'code to write ltw file to disk ================================
                'LogFile.Write(strLTWOutput)
                'CreateOutputFile(strLTWOutput)

                If boolMSHExists Then ' 20140324 only create n NVP file if the file has an MSH segment
                    CreateOutputFile(strLTWOutput)
                End If
                theFile.Delete()
                '===============================================================
                'End If
            Next

            'LogFile.Close()
        Catch ex As Exception

            writeTolog(ex.Message, 1)

            If theFile.Exists Then
                theFile.Delete()
            End If

            Exit Sub
        End Try
    End Sub
    Public Sub CreateOutputFile(ByVal strLTWOutput As String)
        'Function to create an HL7 output file
        Try
            Dim line As String = ""
            Dim objTStreamCounter As Object
            Dim intCounter As Integer = 0

            Dim filename As String
            Dim objTStreamOutput As Object

            'If the file does not exist, create it.
            If Not File.Exists(strOutputDirectory & "counter.txt") Then
                objTStreamCounter = File.CreateText(strOutputDirectory & "counter.txt")
                objTStreamCounter.WriteLine("000")
                objTStreamCounter.Close()
            End If

            'read the present file number for counter.Txt. convert it to an integer and increment it.
            objTStreamCounter = New StreamReader(strOutputDirectory & "counter.txt")

            line = objTStreamCounter.readline
            intCounter = CInt(line)
            intCounter = intCounter + 1
            If intCounter >= 100000 Then intCounter = 0
            objTStreamCounter.Close()

            'write the LTW file to strOutputDirectory a new file is created
            'If strOutputSubDirectory <> "" Then
            'filename = strOutputDirectory & "\" & strOutputSubDirectory & "\LTW." & padleft(Str(intCounter), 3)
            'Else

            '20121013 - changed to create NVP Files
            'filename = strOutputDirectory & "\LTW." & padleft(Str(intCounter), 3)
            filename = strOutputDirectory & "\NVP." & padleft(Str(intCounter), 3)
            'End If

            objTStreamOutput = File.AppendText(filename)
            objTStreamOutput.Write(strLTWOutput)
            objTStreamOutput.Close()

            'update the counter file
            objTStreamCounter = New StreamWriter(strOutputDirectory & "counter.txt")
            objTStreamCounter.WriteLine(padleft(Str(intCounter), 3))
            objTStreamCounter.Close()
        Finally
        End Try
    End Sub


    Public Function padleft(ByRef inputStr As String, ByRef strLength As Short) As String
        'pad an input string with zeros based on desired strLength
        Try
            Dim varLength As Short
            Dim strOutput As String
            Dim i As Short


            strOutput = ""
            varLength = Len(Trim(inputStr))
            For i = 1 To ((strLength - varLength))
                strOutput = strOutput & "0"
            Next
            strOutput = strOutput & Trim(inputStr)
            padleft = strOutput

        Finally
        End Try
    End Function
    Public Sub CreateHashTable()
        Try
            Dim s As String
            Dim segID As String = ""
            Dim segDescription As String = ""
            Dim delimStr As String = "="
            Dim strLine As String = ""
            Dim delimiter As Char() = delimStr.ToCharArray()
            Using sr As StreamReader = New StreamReader(strMapperFile)
                Dim line As String
                ' Read and display the lines from the file until the end 
                ' of the file is reached.
                Do

                    line = sr.ReadLine()

                    If line <> "" Then
                        Dim myArray As String() = Nothing
                        myArray = line.Split(delimiter)
                        segID = myArray(0)
                        segDescription = myArray(1)
                        If myHT.ContainsKey(segID) = False Then
                            myHT.Add(segID, segDescription)
                        End If
                    End If
                    'Next
                Loop Until line Is Nothing
                sr.Close()
            End Using
            'TextBox1.AppendText(myHT.Item("MSH_3"))
        Finally
        End Try
    End Sub
    Public Sub CreateOutputSubDirectory(ByVal s As String)
        Try
            If s <> "" Then
                Dim di As DirectoryInfo = New DirectoryInfo(strOutputDirectory & "\" & s)
                If di.Exists Then
                    strOutputSubDirectory = s
                Else
                    di.Create()
                    strOutputSubDirectory = s
                End If
            End If
        Finally
        End Try

    End Sub
    Public Function BuildSegCounter(ByRef theCounter As Integer) As String
        BuildSegCounter = ""
        Try
            If theCounter <= 1 Then
                BuildSegCounter = ""
            End If
            If theCounter > 1 And theCounter < 10 Then
                BuildSegCounter = "_000" & theCounter
            End If

            If theCounter >= 10 And theCounter < 100 Then
                BuildSegCounter = "_00" & theCounter
            End If

            If theCounter >= 100 And theCounter < 1000 Then
                BuildSegCounter = "_0" & theCounter
            End If

            If theCounter >= 1000 Then
                BuildSegCounter = "_" & theCounter
            End If
        Catch ex As Exception
        End Try
    End Function
    Public Sub writeTolog(ByVal strMsg As String, ByVal eventType As Integer)
        '20140205 - use a text file to log errors instead of the event log
        Dim file As System.IO.StreamWriter
        Dim tempLogFileName As String = strLogDirectory & "ITWMapper_log.txt"
        file = My.Computer.FileSystem.OpenTextFileWriter(tempLogFileName, True)
        file.WriteLine(DateTime.Now & " : " & strMsg)
        file.Close()
    End Sub

    Public Sub writeToLog2(ByVal logText As String, ByVal eventType As Integer)
        Dim myLog As New EventLog()
        Try
            ' check for the existence of the log that the user wants to create.
            ' Create the source, if it does not already exist.
            If Not EventLog.SourceExists("STAR_ITWMapper") Then
                EventLog.CreateEventSource("STAR_ITWMapper", "STAR_ITWMapper")
            End If

            ' Create an EventLog instance and assign its source.

            myLog.Source = "STAR_ITWMapper"

            ' Write an informational entry to the event log.
            If eventType = 1 Then
                myLog.WriteEntry(logText, EventLogEntryType.Error, 1)
            ElseIf eventType = 2 Then
                myLog.WriteEntry(logText, EventLogEntryType.Warning, 2)
            ElseIf eventType = 3 Then
                myLog.WriteEntry(logText, EventLogEntryType.Information, 3)
            End If


        Finally
            myLog.Close()
        End Try
    End Sub
End Module
