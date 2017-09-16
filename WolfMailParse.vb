
Namespace POP3
    ''' <summary>
    ''' text content
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure SContent
        ''' <summary>
        ''' holds content as a string.
        ''' </summary>
        ''' <remarks></remarks>
        Public content As String
        ''' <summary>
        ''' Holds headers associated with this particular content.
        ''' Can be empty in some cases, most notibly anytime a message is not of a multipart/... type.
        ''' </summary>
        ''' <remarks></remarks>
        Public headers As Dictionary(Of String, String)
    End Structure
    ''' <summary>
    ''' Binary content
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure BContent
        ''' <summary>
        ''' holds content in binary representation.
        ''' </summary>
        ''' <remarks></remarks>
        Public content As Byte()
        ''' <summary>
        ''' Holds headers associated with this particular content.
        ''' Can be empty in some cases, most notibly anytime a message is not of a multipart/... type.
        ''' </summary>
        ''' <remarks></remarks>
        Public headers As Dictionary(Of String, String)
        ''' <summary>
        ''' helper function. Can use to save content.
        ''' designed specifically for attachments.
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub SaveAs()
            'Content-Disposition: attachment; filename="PizzaCake2.JPG"; size=10310;
            'creation-date="Sat, 16 Sep 2017 05:38:25 GMT";
            'modification-date="Sat, 16 Sep 2017 05:38:25 GMT"

            'Content-Type: image/jpeg; name="PizzaCake2.JPG"

            'Content-Description: PizzaCake2.JPG

            '3 possible solutions.

            Dim fname = ""
            If headers.ContainsKey("Content-Type") Then
                Dim TXT = LCase(headers("Content-Type"))
                If InStr(TXT, "name") Then
                    Dim fn() = Split(TXT, "name=")
                    fname = Replace(fn(1), """", "")
                End If
            ElseIf headers.ContainsKey("Content-Description") Then
                Dim TXT = LCase(headers("Content-Type"))
                If InStr(TXT, "name") Then
                    fname = Replace(TXT, """", "")
                End If
            ElseIf headers.ContainsKey("Content-Disposition") Then
                Dim TXT = LCase(headers("Content-Type"))
                If InStr(TXT, "name") Then
                    Dim tx() = Split(TXT, ";")
                    For Each Cx In tx
                        If InStr(Cx, ("filename")) Then
                            Dim fn() = Split(Cx, "filename=")
                            fname = Replace(fn(1), """", "")
                            Exit For
                        End If
                    Next
                End If
                Dim SW = New System.Windows.Forms.SaveFileDialog
                SW.FileName = fname
                Dim ret = SW.ShowDialog()
                If ret = Windows.Forms.DialogResult.OK Or ret = Windows.Forms.DialogResult.Yes Then
                    My.Computer.FileSystem.WriteAllBytes(SW.FileName, content, False)
                End If
            End If
        End Sub
    End Structure
    ''' <summary>
    ''' Enum defining different content stoage types (text/binary)
    ''' Again, must be distinguished from the Content-Type header.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ContentStoreType
        text = 0
        binary = 1
    End Enum
    ''' <summary>
    ''' used to represent text or binary content.
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure Content
        ''' <summary>
        ''' Indicates type of content (text/binary)
        ''' not to be confused with content-type header (which should be accessed from 'headers')
        ''' </summary>
        ''' <remarks></remarks>
        Public type As ContentStoreType
        ''' <summary>
        ''' holds the content (either a SContent object for string (text) content, or BContent object for binary content)
        ''' </summary>
        ''' <remarks></remarks>
        Public obj As Object
        ''' <summary>
        ''' Helper function. Applies type conversion to return as a SContent object (if text type)
        ''' </summary>
        ''' <returns>Object if content is text, nothing for binary</returns>
        ''' <remarks></remarks>
        Public Function GetTextContent() As SContent
            If TypeOf obj Is SContent Then
                Return CType(obj, SContent)
            Else
                Return Nothing
            End If
        End Function
        ''' <summary>
        ''' Helper function. Applies type conversion to return as a BContent object (if binary type)
        ''' </summary>
        ''' <returns>Object if content is binary, nothing for text</returns>
        ''' <remarks></remarks>
        Public Function GetBinContent() As BContent
            If TypeOf obj Is BContent Then
                Return CType(obj, BContent)
            Else
                Return Nothing
            End If
        End Function
    End Structure
    ''' <summary>
    ''' Used to parse the recieved message.
    ''' complete redesign as the parser included in the original material was lacking in this aspect.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class MailMessage

        Private m_src As String
        ''' <summary>
        ''' contains headers associated with the email.
        ''' headers for content (when content is sent as a multipart/...) is stored separately
        ''' </summary>
        ''' <remarks></remarks>
        Public headers As Dictionary(Of String, String)
        ''' <summary>
        ''' Contains all available contents associated with the email. 
        ''' for multipart/... types, this will contain multiple contents (along with any headers associated with each)
        ''' for single content messages, only one content exists. In this case, all headers are contained in the MailMessage.headers dictionary.
        ''' </summary>
        ''' <remarks></remarks>
        Public contents As List(Of Content)
        'some duplicate headers may exist. Will take care of this by catentation.
        ''' <summary>
        ''' Create a new instance (passing in message source)
        ''' </summary>
        ''' <param name="source">message source</param>
        ''' <remarks>More things can go wrong. Everything seems to work fine so far, though more testing may be
        ''' needed to verify errors happen when supposed to and check for possible bugs.</remarks>
        Public Sub New(ByVal source As String)
            m_src = source
            'to hold headers
            headers = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            'to hold contents
            contents = New List(Of Content)
            'parse out all headers first
            ParseHeaders(m_src, headers)

            Dim CT = headers("Content-Type"), boundary As String = ""
            If LCase(CT) Like "*multipart/*" Then
                Dim PS = Split(CT, "boundary=")
                boundary = ("--" & Replace(PS(1), """", ""))
                ParseContent(m_src, boundary, headers)
            Else
                Dim R = Split(m_src, vbCrLf & vbCrLf, 2)
                Dim SCont = New SContent With {.content = R(1), .headers = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)}
                Dim cont = New Content With {.obj = SCont, .type = ContentStoreType.text}
                contents.Add(cont)
            End If

        End Sub
        ''' <summary>
        ''' helper function. Gets 'From' address from headers dictionary.
        ''' </summary>
        ''' <returns>string</returns>
        ''' <remarks></remarks>
        Public Function GetFrom() As String
            If headers.ContainsKey("From") Then
                Return headers("From")
            Else
                Return ""
            End If
        End Function
        ''' <summary>
        ''' helper function. Gets 'To' address from headers dictionary.
        ''' </summary>
        ''' <returns>string</returns>
        ''' <remarks></remarks>
        Public Function GetTo() As String
            If headers.ContainsKey("To") Then
                Return headers("To")
            Else
                Return ""
            End If
        End Function
        ''' <summary>
        ''' helper function. Gets subject from headers dictionary.
        ''' </summary>
        ''' <returns>string</returns>
        ''' <remarks></remarks>
        Public Function GetSubject() As String
            If headers.ContainsKey("Subject") Then
                Return headers("Subject")
            Else
                Return ""
            End If
        End Function
        ''' <summary>
        ''' helper function. Gets 'Thread Topic' from headers dictionary.
        ''' </summary>
        ''' <returns>string</returns>
        ''' <remarks></remarks>
        Public Function GetThreadTopic() As String
            If headers.ContainsKey("Thread-Topic") Then
                Return headers("Thread-Topic")
            Else
                Return ""
            End If
        End Function
        ''' <summary>
        ''' helper function. Gets Date from headers dictionary.
        ''' </summary>
        ''' <returns>string</returns>
        ''' <remarks></remarks>
        Public Function GetDateRecieved() As String

            If headers.ContainsKey("Date") Then
                Return headers("Date")
            Else
                Return ""
            End If
        End Function
        ''' <summary>
        ''' helper function. Gets 'Recieved' header from headers dictionary.
        ''' </summary>
        ''' <returns>string</returns>
        ''' <remarks></remarks>
        Public Function GetReceivedPath() As String
            If headers.ContainsKey("Received") Then
                Return headers("Received")
            Else
                Return ""
            End If
        End Function
        Private Function QPS(ByVal inp As String) As String
            Dim X As IO.StringReader = New IO.StringReader(inp), Y As Text.StringBuilder = New Text.StringBuilder
            Dim Z = X.Peek

            While X.Peek >= 0
                Dim A As Char = Chr(X.Read)
                'if char is '=' then we found a code. All real '=' are '=3D'.
                If A = "=" Then
                    Dim B(1) As Char
                    X.ReadBlock(B, 0, 2)
                    Dim E As Integer
                    If B <> vbCrLf Then 'a trailing = is a soft break, otherwise we should get =XX (x is hex).
                        Try
                            E = Convert.ToInt32(B, 16)
                        Catch ex As Exception
                            E = Asc(" ") 'just pad a space in error scenario.
                        End Try
                        Y.Append(Chr(E))
                    Else
                        'do nothing. =[CRLF] is a 'soft' break to cut a line into shorter sections.
                    End If
                Else
                    Y.Append(A)
                End If
            End While
            Return Y.ToString
        End Function
        Private Sub ParseContent(ByVal src As String, ByVal boundary As String, ByVal hdr As Dictionary(Of String, String))
            Dim SRCS() As String = Split(src, vbCrLf & vbCrLf, 2)
            If boundary = "" Then
                'parse the (unlikely) case of single content.
                'possibly text message, unlikely anything non text.
                Dim CT As Content
                Dim TC As SContent
                CT.type = ContentStoreType.text
                'check if it comes as 'quoted-printable'
                If hdr.ContainsKey("Content-Transfer-Encoding") AndAlso LCase(headers("Content-Transfer-Encoding")) Like "quoted-printable" Then
                    TC.content = QPS(SRCS(0))
                    '" quoted-printable" content type replaces some chars with codes in form of =XX. (X is hexadecimal, always 3 digits)
                    ' To fix, we find all cases, then get char from hex value 
                Else
                    TC.content = SRCS(0)
                End If
                TC.headers = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                CT.obj = TC
                contents.Add(CT)
            Else

                Dim srcsx() As String = Split(SRCS(1), boundary)
                For Each cnt In srcsx
                    'sometimes there is just an empty content. Ignore it.
                    If Replace(Trim(cnt), vbCrLf, "") = "--" Or Trim(cnt) = "" Then
                        Continue For
                    End If
                    'There is sometimes a new line after the boundary, sometimes a '--'
                    If Left(cnt, 2) = vbCrLf Then cnt = Mid(cnt, 3)
                    Dim hdr2 = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                    ParseHeaders(cnt, hdr2)
                    'just like a full document, each individual section is of format headers then empty line then content.
                    Dim parts() = Split(cnt, vbCrLf & vbCrLf, 2)

                    If hdr2.ContainsKey("Content-Type") Then
                        Dim CT = hdr2("Content-Type")
                        If LCase(CT) Like "*text*" Then
                            ParseTextContent(parts(1), hdr2)
                        ElseIf LCase(CT) Like "*multipart/*" Then
                            Dim boundary2 As String
                            Dim PS = Split(CT, "boundary=")
                            boundary2 = ("--" & Replace(PS(1), """", ""))
                            ParseContent(parts(1), boundary2, hdr2)
                        Else
                            ParseBinContent(parts(1), hdr2)
                        End If
                    Else
                        'no headers most likely.
                        Dim TC As SContent
                        Dim tct As Content
                        TC.content = parts(0) & vbNewLine & parts(1)
                        TC.headers = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                        tct.obj = TC
                        tct.type = ContentStoreType.text
                        contents.Add(tct)
                    End If
                Next
            End If
        End Sub
        Private Sub ParseBinContent(ByVal src As String, ByVal hdr As Dictionary(Of String, String))
            Dim cont As Content
            Dim bcont As BContent
            Dim B() As Byte
            If hdr.ContainsKey("Content-Transfer-Encoding") Then
                Dim CTE = hdr("Content-Transfer-Encoding")
                If LCase(CTE) Like "*base64*" Then
                    B = Convert.FromBase64String(src)
                ElseIf LCase(CTE) Like "*quoted-printable*" Then
                    B = Text.Encoding.UTF8.GetBytes(QPS(src))
                Else
                    B = Text.Encoding.UTF8.GetBytes((src))
                End If
            Else
                B = Text.Encoding.UTF8.GetBytes((src))
            End If
            bcont.content = B
            bcont.headers = hdr
            cont.obj = bcont
            cont.type = ContentStoreType.binary
            contents.Add(cont)
        End Sub
        Private Sub ParseTextContent(ByVal src As String, ByVal hdr As Dictionary(Of String, String))
            Dim cont As Content
            Dim scont As SContent
            Dim out As String
            'Content-Transfer-Encoding: base64
            'Content-Transfer-Encoding: quoted-printable
            If hdr.ContainsKey("Content-Transfer-Encoding") Then
                Dim CTE = hdr("Content-Transfer-Encoding")
                If LCase(CTE) Like "*quoted-printable*" Then
                    out = (QPS(src))
                ElseIf LCase(CTE) Like "*base64" Then
                    Dim code = System.Text.Encoding.UTF8
                    out = code.GetString(Convert.FromBase64String(src))
                Else
                    out = src
                End If
            Else
                out = src
            End If
            scont.content = out
            scont.headers = hdr
            cont.obj = scont
            cont.type = ContentStoreType.text
            contents.Add(cont)
        End Sub
        Private Sub ParseHeaders(ByVal src As String, ByRef heads As Dictionary(Of String, String))
            Dim SR = New IO.StringReader(src)
            Dim NL As String, name As String = "", value As String = ""
            NL = SR.ReadLine()
            'loop through all lines untill we reach the first empty one.
            'just like http, a empty line indicates end of the header.
            While (NL <> "")
                'we now have support for multiline headers. more complication.
                'so we read lines until we find the tale-tale colon(:), appending them
                'to the last headers value.
                If InStr(NL, ":") And NL(0) <> vbTab And NL(0) <> " " Then
                    If name <> "" Then
                        If heads.ContainsKey(name) Then
                            'also some headers can occur more than once.
                            'this seems limited to the 'recieved' header.
                            'so just cat together any similar named headers.
                            'add empty line, so the different values will be
                            'easy to separate as empty lines are not valid
                            'in the header values (since they indicate end of header)
                            heads(name) &= vbNewLine & vbNewLine & value
                        Else
                            heads.Add(name, value)
                        End If
                        ' name = ""
                        ' value = ""
                    End If
                    Dim S2 = Split(NL, ":", 2)
                    name = S2(0)
                    value = S2(1)
                Else
                    value &= vbNewLine & NL
                End If
                NL = SR.ReadLine
            End While
            If name <> "" Then
                If heads.ContainsKey(name) Then
                    heads(name) &= vbNewLine & value
                Else
                    heads.Add(name, value)
                End If
            End If
        End Sub
    End Class
End Namespace