# POP3.net
This is a POP3 client written entirely in Visual Basic .NET that also includes a message parser to parse recieved emails, including attachments.

The POP3 client portion is designed to conform with rfc1939 and supports most optional commands (except APOP command), and supports both insecure (open) and secure (SSL) connections.

The message parser is designed to deal with mulipart messages as well as single content. It parses out all main headers into a dictionary. It also parses the content section. For multipart messages, the list contains each part along with its associated headers. For content-transfer-encodings: base64, quoted-printable, 7-bit, 8-bit, binary all should be supported.

Only two code files need to be added to a project to make it work, WolfPop3.vb and WolfMailParse.vb. These only rely on tcpclient (system.net.sockets) and sslstream (system.net.security).
