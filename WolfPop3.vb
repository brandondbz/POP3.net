Imports System.Net.Sockets
Imports System.Net.Security

'todo: Test it more fully. find bugs, fix bugs.
Namespace POP3
    Namespace status
        ''' <summary>
        ''' Used for STAT command
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure Status
            Public NumMessages As Integer
            Public TotalOctates As Int64
            Public other As String
        End Structure
        ''' <summary>
        ''' used for LIST command
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure MsgStatus
            Public msgNum As Integer
            Public msgSize As Integer
        End Structure
    End Namespace
    ''' <summary>
    ''' The main POP3 class.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class WolfPop3
        'note: when a message contains a single dot on a line 
        '.
        'then it pads with a second
        '..
        'this is because when receiving a message using the RETR command
        'a single dot on the line indicates the full message has been recieved.
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <remarks>useless mainly. TODO: repalce all usage with vbCrLf</remarks>
        Public Const term = vbCrLf
        Private soc As TcpClient
        Private ssl As SslStream
        Private stream As IO.Stream
        Private SRead As IO.StreamReader
        Private SWrite As IO.StreamWriter
        Private usessl_ As Boolean
        Public LoggedOn As Boolean = False
        ''' <summary>
        ''' initiates new POP3 client
        ''' </summary>
        ''' <param name="Server">Server Name</param>
        ''' <param name="USR">User Name</param>
        ''' <param name="PWD">Password</param>
        ''' <param name="UseSSL">True to use SSL, false otherwise</param>
        ''' <param name="port">Port to use.</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal Server As String, ByVal USR As String, ByVal PWD As String, ByVal UseSSL As Boolean, ByVal port As Integer)
            Try
                'open new socket
                soc = New TcpClient(Server, port)
                'Authenticate using SSL
                usessl_ = UseSSL
                If UseSSL Then 'realllllly recomended 
                    ssl = New SslStream(soc.GetStream)
                    ssl.AuthenticateAsClient(Server)
                    'Create reader/writer for simplicity.
                    SRead = New IO.StreamReader(ssl)
                    SWrite = New IO.StreamWriter(ssl)
                Else
                    stream = soc.GetStream
                    SRead = New IO.StreamReader(stream)
                    SWrite = New IO.StreamWriter(stream)
                End If
                'must be CrLf never Cr or Lf or LfCr.
                SWrite.NewLine = vbCrLf
                'Wait for server to start
                Dim LN As String = SRead.ReadLine()
                If LN Like "*+OK*" Then
                    'server up and running
                ElseIf LN Like "*-ERR*" Then
                    Throw New Exception(Replace(LN, "-ERR", "Error:"))
                Else
                    Throw New Exception("Error: Invalid pop3 start responce")
                End If
                'give user name
                SWrite.WriteLine("USER " & USR)
                SWrite.Flush()

                LN = SRead.ReadLine()
                If LN Like "*+OK*" Then
                    'USER OK (maybe. some server may wait for password to check.)
                ElseIf LN Like "*-ERR*" Then
                    Throw New Exception(Replace(LN, "-ERR", "Error:"))
                Else
                    Throw New Exception("Error: Invalid pop3 start responce")
                End If

                SWrite.WriteLine("PASS " & PWD)
                SWrite.Flush()

                LN = SRead.ReadLine()
                If LN Like "*+OK*" Then
                    'PASS OK (maybe. some server may wait for password to check.)
                ElseIf LN Like "*-ERR*" Then
                    Throw New Exception(Replace(LN, "-ERR", "Error:"))
                Else
                    Throw New Exception("Error: Invalid pop3 start responce")
                End If
                'so the user of the library can see if we successfully logged on.
                LoggedOn = True

            Catch ex As Exception
                If ex.Message = "Error: Invalid pop3 start responce" Then
                    Throw ex
                Else
                    Throw New Exception("Error: Error occured starting client.", ex)
                End If
            End Try
        End Sub
        'no-op does nothing, but does return positive.
        ''' <summary>
        ''' Performs a no-operation.
        ''' Usefull to keep logged in
        ''' </summary>
        ''' <returns>True unless an error occurs.</returns>
        ''' <remarks></remarks>
        Public Function Poll() As Boolean
            'no operation command. keeps connection alive.
            SWrite.WriteLine("NOOP ")
            SWrite.Flush()
            'check responce. Return true for ok, false for err.
            Dim LN = SRead.ReadLine()
            If LN Like "*+OK*" Then
                Return True
            ElseIf LN Like "*-ERR*" Then
                Return False
            Else
                Throw New Exception("Error: Invalid pop3 start responce")
            End If
        End Function

        'Get stats
        ''' <summary>
        ''' Returns number of emails and total estimated storage used by them
        ''' </summary>
        ''' <returns>status object</returns>
        ''' <remarks></remarks>
        Public Function GetStats() As status.Status
            SWrite.WriteLine("STAT ")
            SWrite.Flush()

            'gets number of messages, and total octets used.
            Dim LN = SRead.ReadLine()
            If LN Like "*+OK*" Then
                Dim parts() = Split(LN, " ")
                Dim rt As status.Status = New status.Status With {.other = ""}
                rt.NumMessages = Val(parts(1))
                If parts.Length > 2 Then
                    rt.TotalOctates = Val(parts(2))
                    If parts.Length > 3 Then
                        rt.other = parts(3)
                    End If
                End If
                Return rt
            ElseIf LN Like "*-ERR*" Then
                Throw New Exception(Replace(LN, "-ERR", "Error:"))
            Else
                Throw New Exception("Error: Invalid pop3 start responce")
            End If
        End Function

        'list. Get size of a single message, or list all messages and their sizes.
        ''' <summary>
        ''' Gets message number and size given a message number.
        ''' </summary>
        ''' <param name="msg">message number</param>
        ''' <returns>returns both the message number and it's size.</returns>
        ''' <remarks></remarks>
        Public Function List(ByVal msg As Integer) As status.MsgStatus
            SWrite.WriteLine("LIST " & msg)
            SWrite.Flush()

            'lists message number (the argument) and message size.
            'error if message does not exist.
            Dim LN = SRead.ReadLine()
            If LN Like "*+OK*" Then
                Dim rt As status.MsgStatus = New status.MsgStatus
                Dim parts = Split(LN, " ")
                rt.msgNum = Val(parts(1))
                rt.msgSize = Val(parts(2))
                Return rt
            ElseIf LN Like "*-ERR*" Then
                Throw New Exception(Replace(LN, "-ERR", "Error:"))
            Else
                Throw New Exception("Error: Invalid pop3 start responce")
            End If
        End Function
        ''' <summary>
        ''' helper. Wraps 'List' function, but only returns size.
        ''' </summary>
        ''' <param name="num">message number</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function MsgSize(ByVal num As Integer) As Integer
            Return List(num).msgSize
        End Function

        'list all messages.
        ''' <summary>
        ''' lists all numbers, giving their message number and size.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function list() As List(Of status.MsgStatus)
            SWrite.WriteLine("LIST ")
            SWrite.Flush()
            'list command withoug argument generates a list of number size pairs.
            Dim LN = SRead.ReadLine()
            If LN Like "*+OK*" Then
                LN = SRead.ReadLine()
                Dim ret = New List(Of status.MsgStatus)
                While (LN <> ".")
                    Dim rt As status.MsgStatus = New status.MsgStatus
                    Dim parts = Split(LN, " ")
                    rt.msgNum = Val(parts(1))
                    rt.msgSize = Val(parts(2))
                    ret.Add(rt)
                    LN = SRead.ReadLine()
                End While
                Return ret
            ElseIf LN Like "*-ERR*" Then
                Throw New Exception(Replace(LN, "-ERR", "Error:"))
            Else
                Throw New Exception("Error: Invalid pop3 start responce")
            End If
        End Function
        Private Function StrOnly(ByVal ip As String, ByVal ch As String) As Boolean
            For i = 0 To ip.Length - 1
                If ip(i) <> ch Then Return False
            Next
            Return True
        End Function
        'gets a message
        ''' <summary>
        ''' Retrieve the specified message. You can use the mailmessage class to parse it out.
        ''' </summary>
        ''' <param name="num">message number</param>
        ''' <returns>The message</returns>
        ''' <remarks></remarks>
        Public Function GetMessage(ByVal num As Integer) As String
            SWrite.WriteLine("RETR " & num)
            SWrite.Flush()
            'Gets a message.
            Dim LN = SRead.ReadLine()
            If LN Like "*+OK*" Then
                Dim b = New Text.StringBuilder
                LN = SRead.ReadLine
                'read all lines untill we reach the end.
                While (LN <> ".")
                    If StrOnly(LN, ".") Then
                        b.Append(Mid(LN, 2) & term)
                    Else
                        b.Append(LN & term)
                    End If
                    LN = SRead.ReadLine()
                End While
                Return b.ToString()
            ElseIf LN Like "*-ERR*" Then
                Throw New Exception(Replace(LN, "-ERR", "Error:"))
            Else
                Throw New Exception("Error: Invalid pop3 start responce")
            End If
        End Function
        'delete message. more accurately, marks message deleted (like recycle bin).
        ''' <summary>
        ''' Marks a message for deletion. Made permenant upon calling 'Quit'
        ''' </summary>
        ''' <param name="num">message number</param>
        ''' <returns>true on success</returns>
        ''' <remarks></remarks>
        Public Function DeleteMessage(ByVal num As Integer) As Boolean
            SWrite.WriteLine("DELE " & num)
            SWrite.Flush()
            'Gets a message.
            Dim LN = SRead.ReadLine()
            If LN Like "*+OK*" Then
                Return True
            ElseIf LN Like "*-ERR*" Then
                Return False
            Else
                Throw New Exception("Error: Invalid pop3 start responce")
            End If
        End Function
        'Reset all marked for deletions. causes all messages to return to status they had when client started.
        ''' <summary>
        ''' Resets all messages that were marked for deletion. Calling this prevents marked messages from being deleted on quit.
        ''' </summary>
        ''' <returns>true on success</returns>
        ''' <remarks></remarks>
        Public Function ResetDeletions() As Boolean
            SWrite.WriteLine("RSET ")
            SWrite.Flush()
            'Gets a message.
            Dim LN = SRead.ReadLine()
            If LN Like "*+OK*" Then
                Return True
            ElseIf LN Like "*-ERR*" Then
                Return False
            Else
                Throw New Exception("Error: Invalid pop3 start responce")
            End If
        End Function
        'Must call quit to finalize deletions. Not calling quit, mean nothing can be deleted. (same as calling reset then quit)
        ''' <summary>
        ''' Call to carry out any marked deletions.
        ''' </summary>
        ''' <returns>true on success</returns>
        ''' <remarks>Calling ResetDeletions then this will have no net effect. 
        ''' Not calling Quit after marking messages for deletions will have the same effect as calling ResetDeletions before calling quit.</remarks>
        Public Function Quit() As Boolean
            SWrite.WriteLine("QUIT ")
            SWrite.Flush()
            'quit. Cause all messages marked for delete to die.
            Dim LN = SRead.ReadLine()
            If LN Like "*+OK*" Then
                Return True
            ElseIf LN Like "*-ERR*" Then
                Return False
            Else
                Throw New Exception("Error: Invalid pop3 start responce")
            End If
        End Function
        'Gets top N lines of message. Why? IDK!
        ''' <summary>
        ''' Gets the first N lines of message num.
        ''' The N lines does not include headers
        ''' </summary>
        ''' <param name="num">message number</param>
        ''' <param name="N"># of lines to return</param>
        ''' <returns>Part of message</returns>
        ''' <remarks></remarks>
        Public Function GetMessageTop(ByVal num As Integer, ByVal N As Integer) As String
            SWrite.WriteLine("TOP " & num & " " & N)
            SWrite.Flush()
            'Gets a message.
            Dim LN = SRead.ReadLine()
            If LN Like "*+OK*" Then
                Dim b = New Text.StringBuilder
                LN = SRead.ReadLine
                'read all lines untill we reach the end.
                While (LN <> ".")
                    If StrOnly(LN, ".") Then
                        b.Append(Mid(LN, 2) & term)
                    Else
                        b.Append(LN & term)
                    End If
                    LN = SRead.ReadLine()
                End While
                Return b.ToString()
            ElseIf LN Like "*-ERR*" Then
                Throw New Exception(Replace(LN, "-ERR", "Error:"))
            Else
                Throw New Exception("Error: Invalid pop3 start responce")
            End If
        End Function
        'Gets message unique identifier. 
        ''' <summary>
        ''' Get the Unique IDentifier (UID) of message num.
        ''' </summary>
        ''' <param name="num">message number</param>
        ''' <returns>message unique identifier</returns>
        ''' <remarks>Not much you can do with the UID.</remarks>
        Public Function MUID(ByVal num As Integer) As String
            SWrite.WriteLine("UIDL " & num)
            SWrite.Flush()
            'gets the message uid.
            Dim LN = SRead.ReadLine()
            If LN Like "*+OK*" Then
                Dim parts = Split(LN)
                Return parts(2)
            ElseIf LN Like "*-ERR*" Then
                Throw New Exception(Replace(LN, "-ERR", "Error:"))
            Else
                Throw New Exception("Error: Invalid pop3 start responce")
            End If

        End Function
        ''' <summary>
        ''' Gets the UID for each message. returns them as a dictionary.
        ''' </summary>
        ''' <returns>Dictionary with Key as message number and values as the UIDs.</returns>
        ''' <remarks></remarks>
        Public Function MUID() As Dictionary(Of Integer, String)
            SWrite.WriteLine("UIDL ")
            SWrite.Flush()
            'gets the message uid.
            Dim LN = SRead.ReadLine()
            If LN Like "*+OK*" Then
                LN = SRead.ReadLine()
                Dim ret = New Dictionary(Of Integer, String)
                While (LN <> ".")
                    Dim parts = Split(LN)
                    Dim id As Integer = Val(parts(0))
                    If Not ret.ContainsKey(id) Then ret.Add(id, parts(1))
                End While
                Return ret
            ElseIf LN Like "*-ERR*" Then
                Throw New Exception(Replace(LN, "-ERR", "Error:"))
            Else
                Throw New Exception("Error: Invalid pop3 start responce")
            End If
        End Function
    End Class
End Namespace