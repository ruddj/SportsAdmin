Option Compare Database
Option Explicit

Global GlobalGenerateHTML As Variant
Global Template As String
Global TemplateSummary As String
Global GroupLvl As Integer
Global HTMLgenerateFinished As Integer
Global BGcolor As String

Dim s As String

Global Const cWhite = "#FFFFFF"
Global Const cGreen = "#33FF33"
Global Const cCream = "#EEE8D2"
Global Const cBlack = "#000000"
Global Const cLightGray = "#E6E6E6"
Global Const cGray = "#CCCCCC"
Global Const cRed = "#FF0000"
Global Const cLightRed = "#FF8080"

Global Const rGroupHeader = 1
Global Const rGroupFooter = 2
Global Const rDetail = 3
Global Const rPageHeader = 4
Global Const rPageFooter = 5


Type HTMarrayType
    Pg As Integer
    GrpName As Variant
    GrpHead As Integer    ' true or false
    row As String
End Type

Function Alignend(Alignment As String)
    Alignend = "</" & Alignment & ">"
End Function

Function AlignStart(Alignment As String)
    AlignStart = "<" & Alignment & ">"
End Function

Sub CellEnd(HTML As String)
    HTML = HTML & "</td>"
End Sub

Sub CellStart(HTML As String, Align As String, Valign As String, vWidth As String, BGcolor As String, ColSpan As Integer)
    
    s = "<td"
    If ColSpan > 1 Then s = s & " COLSPAN=""" & ColSpan & """"
    If Align <> "" Then s = s & " ALIGN=" & Align
    If Valign <> "" Then s = s & " VALIGN=" & Valign
    If vWidth <> "" Then s = s & " Width=""" & vWidth & """"
    If BGcolor <> "" Then s = s & " bgcolor=""" & BGcolor & """"
    s = s & ">"
    
    HTML = HTML & s

End Sub

Sub CreateHTMLfile(ByVal FileName As String, ByVal TemplateFilename As String, HTML As String, Prev As String, Nex As String, Title As String, Head As String, Optional repClass As String = "report")
    
    Dim HTMLFileLocation, FileLocation, L As String, TemplateFile As String
    Dim HTMLinserted As Integer, Continue As Integer, i As Integer, tFile As Variant, oFile As Variant
    
    HTMLFileLocation = DLookup("[HTMLlocation]", "MiscHTML")
    FileLocation = HTMLFileLocation & "\" & FileName
    
    tFile = FreeFile
    Open TemplateFilename For Input As #99
    oFile = FreeFile
    Open FileLocation For Output As #1 ' Open file for output.
    
    Continue = True
    Do While Not EOF(99)
        HTMLinserted = False
        Input #99, L
        i = 1
        Do While (i <= Len(L))
            If Mid(L, i, 1) = "{" Then
                Select Case UCase(Mid(L, i, 6))
                    Case "{HTML}"
                        Print #1, HTML;
                        i = i + 5
                    Case "{PREV}"
                        Print #1, Prev;
                        i = i + 5
                    Case "{NEXT}"
                        Print #1, Nex;
                        i = i + 5
                    Case "{HEAD}"
                        Print #1, Head;
                        i = i + 5
                    Case "{TITL}"
                        Print #1, Title;
                        i = i + 5
                    Case "{REPN}"
                        Print #1, repClass;
                        i = i + 5
                    Case Else
                        Print #1, "{";
                End Select
            Else
                Print #1, Mid(L, i, 1);
            End If
        i = i + 1
        Loop
        
        Print #1, vbNewLine;
        
    Loop
    
    Close
    
End Sub

 Function Heading(Level As Integer, T As Variant, Indent As Integer)
    Dim i As Integer
    
    s = ""
    s = "<h" & Trim(str(Level)) & ">"
    For i = 1 To Indent
        s = s & "&nbsp;"
    Next
    s = s & T & "</h" & Trim(str(Level)) & ">" & vbNewLine
    Heading = s
    
End Function

 Function HeadingEnd(Level As Integer)
    HeadingEnd = "</h" & Trim(str(Level)) & ">" & vbNewLine
End Function

Function HeadingStart(Level As Integer)
    HeadingStart = "<h" & Trim(str(Level)) & ">"
End Function

Function HTMLend()
    HTMLend = "</body>" & vbNewLine & "</html>"
End Function

Function HTMLStart(Title As String, Author As String)

    s = "<html>" & vbNewLine & "<head>" & vbNewLine
    s = s & "<meta HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=iso-8859-1"">" & vbNewLine
    s = s & "<meta NAME=""Author"" CONTENT=""" & Author & """>" & vbNewLine
    s = s & "<title>" & Title & "</title>" & vbNewLine
    s = s & "</head>" & vbNewLine
    s = s & "<body>" & vbNewLine

    HTMLStart = s
    
End Function

Function image(Source As String, Alternate As String)
    s = "<IMG SRC=""" & Source & """"
    If Alternate <> "" Then s = s & " ALT=""" & Alternate & """"
    image = s
    
End Function

Function Indent(Count As Integer)
    Dim i As Integer
    s = ""
    For i = 1 To Count
        s = s & "<BLOCKQUOTE>"
    Next i
    Indent = s
End Function

Function Link(LinkSource As String, T As String)
    Link = "<a href=""" & LinkSource & """>" & T & "</a>"
End Function

Function LinkEnd()
    LinkEnd = "</a>"
End Function

Function LinkStart(Link As String)
    LinkStart = "<a href=""" & Link & """>"
End Function

Function ParaStart()
    ParaStart = "<P>"
End Function

Function ParaEnd()
    ParaEnd = "</P>"
End Function

Sub Paragraph(HTML As String, Count As Integer)
    Dim i As Integer
    s = ""
    For i = 1 To Count
        s = s & "<br>"
    Next
    HTML = HTML & s
        
End Sub

Sub RowEnd(HTML As String)
    HTML = HTML & "</tr>" & vbNewLine
End Sub

Sub RowStart(HTML As String, Optional sClass As String)
    s = "<tr"
    If sClass <> "" Then s = s & " class=""" & sClass & """ "
    s = s & ">"
    HTML = HTML & s
End Sub
Sub CellHStart(HTML As String)
    HTML = HTML & "<th>"
End Sub
Sub CellHEnd(HTML As String)
    HTML = HTML & "</th>"
End Sub


Sub SpaceIndent(HTML As String, Count As Integer)
    Dim i As Integer
    s = ""
    For i = 1 To Count
        s = s & "&nbsp;"
    Next
    HTML = HTML & s
End Sub

Sub TableEnd(HTML As String)

    HTML = HTML & "</table>" & vbNewLine
    
End Sub

Sub TableStart(HTML As String, vWidth As String, Height As String, BGcolor As String, Caption As String, Border As Integer)
    s = "<table"
    If vWidth <> "" Then s = s & " Width=""" & vWidth & """ "
    If Height <> "" Then s = s & " HEIGHT=""" & Height & """ "
    If BGcolor <> "" Then
        s = s & " bgcolor=""#" & BGcolor & """"
    Else
        s = s & " bgcolor=""" & cWhite & """"
    End If
    s = s & " BORDER=" & Border
    s = s & ">"
    If Caption <> "" Then s = s & "<CAPTION>" & Caption & "</CAPTION>"

    HTML = HTML & s & vbNewLine

End Sub

Sub TableOpen(HTML As String, sClass As String)
    s = "<table"
    If sClass <> "" Then s = s & " class=""" & sClass & """ "
    s = s & ">"
    HTML = HTML & s & vbNewLine

End Sub

Sub ListOpen(HTML As String, Optional sClass As String)
    s = "<ul"
    If sClass <> "" Then s = s & " class=""" & sClass & """ "
    s = s & ">"
    HTML = HTML & s & vbNewLine

End Sub

Sub ListClose(HTML As String)
    s = "</ul>"
    HTML = HTML & s & vbNewLine
End Sub
Sub ListItem(HTML As String, strText As String, Optional sClass As String = "")
    s = "<li"
    If sClass <> "" Then s = s & " class=""" & sClass & """ "
    s = s & ">" & strText & "</li>"
    HTML = HTML & s & vbNewLine
End Sub

Sub Cell(HTML As String, strText As String, Optional sClass As String = "")
    s = "<td"
    If sClass <> "" Then s = s & " class=""" & sClass & """ "
    s = s & ">" & strText & "</td>"
    HTML = HTML & s
End Sub
Sub CellHead(HTML As String, strText As String, Optional sClass As String = "")
    s = "<th"
    If sClass <> "" Then s = s & " class=""" & sClass & """ "
    s = s & ">" & strText & "</th>"
    HTML = HTML & s
End Sub

Sub TableHeadOpen(HTML As String, sClass As String)
    s = "<thead"
    If sClass <> "" Then s = s & " class=""" & sClass & """ "
    s = s & ">"
    HTML = HTML & s & vbNewLine
End Sub
Sub TableHeadEnd(HTML As String)
    s = "</thead>"
    HTML = HTML & s & vbNewLine
End Sub
Sub TableBodyOpen(HTML As String, sClass As String)
    s = "<tbody"
    If sClass <> "" Then s = s & " class=""" & sClass & """ "
    s = s & ">"
    HTML = HTML & s & vbNewLine
End Sub
Sub TableBodyEnd(HTML As String)
    s = "</tbody>"
    HTML = HTML & s & vbNewLine
End Sub

Sub DivOpen(HTML As String, sClass As String, Optional sID As String = "")
    s = "<div"
    If sID <> "" Then s = s & " id=""" & sID & """ "
    If sClass <> "" Then s = s & " class=""" & sClass & """ "
    s = s & ">"
    HTML = HTML & s & vbNewLine

End Sub
Sub DivClose(HTML As String)

    HTML = HTML & "</div>" & vbNewLine
    
End Sub

Sub Text(HTML As String, Style As String, StyleEnd As String, ByVal T As String)
    s = ""
    If Style <> "" Then s = Style
    s = s & T
    If StyleEnd <> "" Then s = s & StyleEnd
    
    HTML = HTML & s
        
End Sub

Function UnIndent(Count As Integer)
    Dim i As Integer
    s = ""
    For i = 1 To Count
        s = s & "</BLOCKQUOTE>"
    Next i
    UnIndent = s

End Function

Function AlphaNumericOnly(strSource As String) As String
    Dim i As Integer
    Dim strResult As String

    For i = 1 To Len(strSource)
        Select Case Asc(Mid(strSource, i, 1))
            Case 48 To 57, 65 To 90, 97 To 122: 'include 32 if you want to include space
                strResult = strResult & Mid(strSource, i, 1)
        End Select
    Next
    AlphaNumericOnly = strResult
End Function

Function AlphaNumericDashOnly(strSource As String) As String
    Dim i As Integer
    Dim strResult As String

    For i = 1 To Len(strSource)
        Select Case Asc(Mid(strSource, i, 1))
            Case 45, 48 To 57, 65 To 90, 97 To 122: 'include 32 if you want to include space
                strResult = strResult & Mid(strSource, i, 1)
        End Select
    Next
    AlphaNumericDashOnly = strResult
End Function