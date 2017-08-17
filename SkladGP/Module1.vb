Module Module1

    Sub Main()
        'On Error GoTo ErrorHandler
        Dim ConnSQL, ts1, sqlstr, fl, path, buf, arr, folder, kup, sdano, d, d1, rs1, sqlstr1, errfl, k
        ConnSQL = CreateObject("ADODB.Connection")
        ConnSQL.ConnectionString = "Provider=SQLOLEDB;Server=srv-otk;Database=otk;Trusted_Connection=yes;Integrated Security=SSPI;Persist Security Info=False"
        Dim fso, fserr
        'Dim dbupd(10000, 2) As String
        Dim pathstring = "c:\sklad\"
        ConnSQL.Open
        fso = CreateObject("Scripting.FileSystemObject")
        fserr = CreateObject("Scripting.FileSystemObject")

        errfl = fserr.OpenTextFile(pathstring & "err\err.txt", 8, True)
        fl = ""
        path = ""
        folder = fso.GetFolder(pathstring & "OUT")
        kup = 0
        sdano = 0
        d = 0
        k = 0
        For Each file In folder.Files
            If Left(file.Name, 3) = "sgp" And Right(file.name, 4) = ".txt" Then
                path = file
                fl = file.name
                Exit For
            End If
        Next
        If Len(fl) < 5 Then
            Console.WriteLine("Файл данных не найден!")

            System.Threading.Thread.Sleep(7000)
            Exit Sub
        End If



        ts1 = fso.OpenTextFile(path, 1, False)
        Do While Not ts1.AtEndOfStream
            buf = ts1.ReadLine
            arr = Split(buf, ";")
            'sqlstr1 = "SELECT [shtr_kod] FROM dbo.[Изделия] WHERE [shtr_kod]=" & arr(2)
            'If ConnSQL.Execute(sqlstr1).EOF = True Then

            'Console.WriteLine(arr(2) & " не существует")
            'errfl.WriteLine(CDate(arr(0) & " " & arr(1)) & vbTab & Now.ToShortTimeString & vbTab & arr(2) & " не существует")
            'd = d + 1
            'Continue Do
            'End If

            If Left(arr(2), 1) = "S" And UBound(arr) = 2 Then
                sqlstr = "DELETE dbo.sklad WHERE nomup=" & Mid(arr(2), 2, 7)
                sdano = sdano + 1
                'dbupd(k, 2) = arr(2)
                ConnSQL.Execute(sqlstr)

            ElseIf Left(arr(3), 1) = "S" And UBound(arr) = 3 Then
                If ConnSQL.execute("SELECT * FROM dbo.sklad WHERE shtr=" & arr(2)).EOF = True Then
                    'sqlstr = "INSERT INTO dbo.t1 (d1,shtr,upak, def) SELECT '" & Cdate (arr(0) &" "&arr(1)) &"'," &arr(2) &", null," & arr(3)
                    'sqlstr = "Update dbo.Изделия SET [DataUp] ='" & CDate(arr(0) & " " & arr(1)) & "', DefUp =" & arr(3) & ", [NomUp]=null WHERE [shtr_kod]=" & arr(2)
                    'kdef = kdef + 1
                    sqlstr = "Insert INTO dbo.sklad ([Datapr], shtr, nomup) SELECT '" & CDate(arr(0) & " " & arr(1)) & "', " & arr(2) & "," & Mid(arr(3), 2, 7)
                    kup = kup + 1
                    ConnSQL.Execute(sqlstr)
                Else
                    d = d + 1
                End If

            End If
            'ConnSQL.execute = sqlstr
            k = k + 1

        Loop
        'sqlstr = "INSERT INTO dbo.LogUpak (data,CountUp,CountBrak) SELECT getdate()," & kup & "," & kdef
        'ConnSQL.execute = sqlstr


        'ConnSQL.Close


        'ConnSQL.Open
        'ConnSQL.BeginTrans
        'For i = 0 To k - 1
        'rs1 = ConnSQL.execute("SELECT * FROM dbo.sklad WHERE shtr=" & dbupd(i, 2))
        'If dbupd(i, 2) > 0 And rs1.EOF = True Then
        'ConnSQL.Execute(dbupd(i, 1))
        'End If
        'MsgBox(d1)
        'Next
        'ConnSQL.CommitTrans
        ConnSQL.Close
        ts1.Close
        errfl.Close
        ts1 = fso.GetFile(path)
        ts1.move(pathstring & "Arhiv\" & fl)

        Console.WriteLine("")
        Console.WriteLine("Принято: " & kup)
        Console.WriteLine("Сдано: " & sdano & " упаковок")
        Console.WriteLine("Пропущено (дубли): " & d)

        System.Threading.Thread.Sleep(7000)

        'Exit Sub

        'ErrorHandler: ' Error-handling routine. 
        '        Select Case Err.Number ' Evaluate error number. 
        '            Case 55 ' "File already open" error. 
        '                Resume
        '            Case Else
        '                ' Handle other situations here... 
        '        End Select
        '        ' Resume execution at same line 

    End Sub



End Module
