Imports System.Text.RegularExpressions


Module MainModule

    'process args	
    Dim IsQuiet As Boolean
    Dim MatchPattern As String
    Dim ReplacePattern As String

    Dim IsPreview As Boolean = False
    Dim IsIgnoreCase As Boolean = False
    Dim IsIncExt As Boolean = False
    Dim IsFileOnly As Boolean = False
    Dim IsDirOnly As Boolean = False

    Dim CurrentPath As String = System.IO.Directory.GetCurrentDirectory()
    Dim options As RegexOptions

    Dim UndoFilename As String = "rren_undo.bat"
    Dim UndoBatObj As Undo

    Sub Main()
        Dim Args() As String = Environment.GetCommandLineArgs
        If Args.Length < 3 Then
            'Print("missing parameters! quit" & vbCrLf)
            PrintHelp()
        End If

        MatchPattern = Args(1)
        ReplacePattern = Args(2)


        For i As Integer = 3 To Args.Length - 1
            Dim em As String = UCase(Args(i))
            Select Case em
                Case "/I"
                    IsIgnoreCase = True
                Case "/P"
                    IsPreview = True
                Case "/E"
                    IsIncExt = True
                Case "/Q"
                    IsQuiet = True
                Case "/F"
                    IsFileOnly = True
                Case "/D"
                    IsDirOnly = True
                Case Else
                    If Left(em, 3) = "/U:" Then
                        UndoFilename = em.Substring(3, em.Length - 3)
                    Else
                        Call PrintHelp()
                    End If
            End Select
        Next



        'create regex obj
        If IsIgnoreCase Then options = RegexOptions.IgnoreCase

        Dim r As String() = System.IO.Directory.GetFiles(CurrentPath)

        'create undo obj
        UndoBatObj = New Undo(UndoFilename)


        If IsDirOnly Then
            ProcessRename(System.IO.Directory.GetDirectories(CurrentPath))
        ElseIf IsFileOnly Then
            ProcessRename(System.IO.Directory.GetFiles(CurrentPath))
        Else
            Dim dirs, files As String()
            files = System.IO.Directory.GetFiles(CurrentPath)
            dirs = System.IO.Directory.GetDirectories(CurrentPath)
            Array.Copy(dirs, files, dirs.Length)
            ProcessRename(files)
        End If


    End Sub

    Sub ProcessRename(ByVal procQueue As String())

        Dim modFilecounter As Integer = 0

        For Each f1 As String In procQueue

            Dim IsModiflied As Boolean = False 'modifiy flag, indicated file name has changed

            Dim OldName As String = System.IO.Path.GetFileNameWithoutExtension(f1)
            Dim ExtName As String = System.IO.Path.GetExtension(f1)
            Dim NewName As String

            If IsIncExt Then
                'also rename .ext
                OldName = OldName & ExtName
                NewName = Regex.Replace(OldName, MatchPattern, ReplacePattern, options)
            Else
                'maintain .ext
                NewName = Regex.Replace(OldName, MatchPattern, ReplacePattern, options)
                OldName = OldName & ExtName
                NewName = NewName & ExtName
            End If


            If LCase(OldName) <> LCase(NewName) Then
                If Not IsPreview Then
                    Try
                        FileSystem.Rename(f1, NewName)
                    Catch e As Exception
                        Print(e.Message)
                    Finally
                    End Try
                    IsModiflied = True
                End If

                'output formating
                Dim tmp As String
                If f1.Length > 30 Then
                    'tmp = Left(f1, 8) & "..." & Right(f1, 19)
                    tmp = Left(f1, 3) & "..." & Right(f1, 24)
                Else
                    tmp = f1.PadRight(30 - f1.Length)
                End If
                Print(tmp & " -> " & NewName)
            End If

            If IsModiflied Then 'write undo batch
                UndoBatObj.writeline("ren """ & NewName & """ """ & OldName & """")
                modFilecounter += 1
            End If

        Next

        If modFilecounter > 0 Then
            Print(modFilecounter & " files renname successful.")
            If Not (IsQuiet Or IsPreview) Then
                Dim objundo As New Undo(UndoFilename)
            End If
        Else
            Print("nothing match, exiting")
        End If

    End Sub



    Private Class Undo

        Dim buff As String
        Dim file As string

        Sub New(filename As String)
            file = filename
        End Sub


        Sub writeline(str As String)
            buff += str & vbCrLf
        End Sub


        Sub savefile()
            Dim batfilestream As IO.StreamWriter

            file = ".\" & file
            If IO.File.Exists(file) Then
                'overwrite question
                Print(file & " already existed.")
                Print("(y = overwrite, others = skip genrating undo bat file, c = cancel)")

                Dim rc As ConsoleKeyInfo = Console.ReadKey(True)
                If rc.Key = ConsoleKey.Y Then
                    'overwrite
                    batfilestream = CreateUndoWriter(file)
                ElseIf rc.Key = ConsoleKey.C Then
                    'cancel
                    End
                Else
                    'skip
                    Exit Sub
                End If
            Else
                batfilestream = CreateUndoWriter(file)
            End If

            batfilestream.WriteLine(buff)
            batfilestream.Close()
        End Sub


        Private Function CreateUndoWriter(ByVal filename As String) As IO.StreamWriter
            'creatnew file or overwrite exist
            CreateUndoWriter = New IO.StreamWriter(filename, False, System.Text.Encoding.Default)
            CreateUndoWriter.AutoFlush = True
        End Function

    End Class

    Private Sub PrintHelp()
        Print("用正则表达式重命名当前路径下的文件及文件夹:")
        Print("rren ""match"" ""replace"" [/i][/p][/q][/u:undo_bat_filename.bat]")
        Print("  match		匹配正则表达式模板")
        Print("  replace	被替换正则表达式模板")
        Print("  /i		ignore case")
        Print("  /p		preview change")
        Print("  /e		included files/directorys extend name")
        Print("  /q		quiet mode (dosn't output undo file and prompt")
        Print("  /d		rename directorys only, excluded /f option")
        Print("  /f		rename files only")

        Print("")
        Print(System.Reflection.Assembly.GetCallingAssembly.GetName.FullName)

        End
    End Sub

    Sub Print(ByVal msg As String)
        If IsQuiet Then Exit Sub
        Console.WriteLine(msg)
    End Sub


End Module
