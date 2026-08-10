Public Class Context
    Private Shared _instance As Context
    Public Shared ReadOnly Property Instance As Context
        Get
            If _instance Is Nothing Then _instance = New Context()
            Return _instance
        End Get
    End Property

    Public iLogitCounter As Integer = 0
    Public bloghelper As Boolean = True
    Public Property csvFilePath As String = "C:\Log\"

    Public Sub LOGIT(ByVal sMess As String, Optional bTrblSht As Boolean = False)
        ' Reference the Singleton Instance
        Dim ctx = Context.Instance

        Try
            ' 1. Get the stack trace from the helper inside the context
            Dim st = Context.ShowStackTrace()

            If Debugger.IsAttached Then
                Debug.WriteLine(sMess)
                Exit Sub
            End If

            Dim stdisplay = String.Join("|", st)

            ' 2. Check the logging flag
            If Not ctx.bloghelper AndAlso Not bTrblSht Then Exit Sub

            ctx.iLogitCounter += 1

            ' 3. FIX: Move the " | " outside the date format brackets
            Dim ds As String = $"{ctx.iLogitCounter.ToString.PadLeft(5)} {DateTime.Now:yyyy-MM-dd HH:mm:ss} | {stdisplay} {sMess}"

            ' 4. Pathing
            Dim logDir As String = IO.Path.Combine(ctx.csvFilePath, "Logs")
            If Not IO.Directory.Exists(logDir) Then IO.Directory.CreateDirectory(logDir)

            ' 5. Filename
            Dim fileName As String = $"{My.Computer.Name}_EmailScreener_{DateTime.Now:yyyyMMdd_HH}.log"
            Dim fullPath As String = IO.Path.Combine(logDir, fileName)

            Using swLog As New IO.StreamWriter(fullPath, True)
                If sMess IsNot Nothing Then
                    ' Split by any newline character
                    For Each line In sMess.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                        swLog.WriteLine(ds)
                    Next
                End If
            End Using

        Catch ex As Exception
            ' Silent fail for logger
        End Try
    End Sub
    Public Shared Function ShowStackTrace() As List(Of String)
        ShowStackTrace = New List(Of String)
        ' Capture the stack trace
        Dim stackTrace As New StackTrace(True) ' True to capture file information

        ' Display each frame in the stack trace
        For Each frame As StackFrame In stackTrace.GetFrames()
            Dim method As System.Reflection.MethodBase = frame.GetMethod()
            Dim lineNumber As Integer = frame.GetFileLineNumber()
            If method.DeclaringType IsNot Nothing Then
                If method.DeclaringType.FullName.ToLower.StartsWith("hughsgolf") Then
                    'logit($"{method.DeclaringType.FullName}.{method.Name} at line {lineNumber}")
                    If method.Name.ToLower.Contains("logit") Or method.Name.ToLower.Contains("showstacktrace") Then Continue For
                    ShowStackTrace.Add($"{method.DeclaringType.Name}.{method.Name}:Line {lineNumber}")
                End If
            End If
        Next
    End Function
End Class
