Attribute VB_Name = "mdlMain"
Option Explicit

Public strAuth As String

Sub CreateJIRATickets()
'Description:____________________________________________________________________________________________________________________________________________________________________________________________________________________
'
'   - Creates a JIRA bug issue for each outlook e-mail marked as an open task and attaches the e-mail to the issue
'

'0.0 Setup

    'Constants
    
    Const DEFAULT_DESCRIPTION As String = "Created by macro."
    Const DEFAULT_PRIORITY As String = "Low"
    
    'Variables

    Dim olApp As New Outlook.Application    'Outlook
    Dim olItem As Object
    Dim olTask As Outlook.TaskItem
    Dim olItems As Outlook.Items
    Dim olTrgFldr As Outlook.Folder
    
    Dim dict As Object  'Scripting
    Dim dictKey As Variant
    
    Dim wsh As Object   'Windows script host
    
    Dim x As Integer
    Dim strSumm As String, strDesc As String, strEpicId As String, strSend As String, strApiUri As String, strIssKey As String, _
      strPath As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set wsh = CreateObject("WScript.Shell")
    
    Set olTrgFldr = olApp.Session.Folders(TEAM_FOLDER).Folders(LCase(Environ("username")))
    Set olItems = olTrgFldr.Items   'All mail in inbox
    
    If olItems.Count = 0 Then GoTo endSubHandler
    
    olItems.Sort "[ReceivedTime]", True
    
    strAuth = EncodeBase64(Environ("username") + ":" + WINDOWS_PASSWORD)
    
'1.0 Loop through inbox to create tickets
    
    For Each olItem In olItems
        
        With olItem
            
            If Trim(TypeName(olItem)) <> "MailItem" Then GoTo nextItem
            If .FlagStatus <> 2 Then GoTo nextItem 'Skip non-task/completed e-mails
            
            'Create summary and description strings
        
            strDesc = DEFAULT_DESCRIPTION
            
            strSumm = .SenderName & " - " & .Subject
            
            strSumm = Replace(strSumm, "#", "")
            strSumm = Replace(strSumm, "\", "")
            strSumm = Replace(strSumm, Chr(9), "")
            strSumm = Replace(strSumm, Chr(10), "")
            strSumm = Replace(strSumm, Chr(34), "")

            'Determine epic id
            
            strEpicId = GetEpicID(.SenderName)
            
            'Send API request
            
            strSend = CreateStrJSON(strSumm, strDesc, strEpicId, DEFAULT_PRIORITY)
            
            Call PostHTTPRequest(API_URI, strSend, strIssKey)
            
            'Download e-mail
            
            strPath = "C:\Users\@user\Downloads\OL JIRA Downloads\"
            strPath = Replace(strPath, "@user", Environ("username"))
            
'            If Dir(strPath) = "" Then MkDir (strPath)
            
            strPath = strPath & "@uniqueid.msg"
            strPath = Replace(strPath, "@uniqueid", Format(Now, "mm_dd_hh_m_ss") & "_" & x)
            
            .SaveAs strPath
            
            'Remove flag
            
            .MarkAsTask (olMarkComplete)
            .Save
            DoEvents
        
        End With
        
        'Add issue number and path to dictionary
        
        dict.Add key:=strIssKey, Item:=strPath
        
        strIssKey = ""
        
nextItem:
        
        'Stop checking after 200 items
        
        x = x + 1
        If x > 200 Then Exit For
        
    Next olItem
    
'2.0 Loop through dictionary, using command terminal to attach e-mails to JIRA issues

    On Error Resume Next
    
        strIssKey = CStr(dict.Keys()(0))  'Check if any keys have been added
        If strIssKey = "" Then
            Err.Clear
            MsgBox Prompt:="No flagged, incomplete e-mails to push to JIRA.", Buttons:=vbInformation
            GoTo endSubHandler
        End If
    
    On Error GoTo 0

    For Each dictKey In dict.Keys
    
        strIssKey = CStr(dictKey)
        strPath = dict(dictKey)
        
        'Create curl REST API command
        
        strApiUri = API_URI & strIssKey & "/attachments/"
    
        strSend = ""
        strSend = "cmd.exe /S /C "
        strSend = strSend & "curl "
        strSend = strSend & "-H ""Authorization: Basic @authorization"" "
        strSend = strSend & "-H ""X-Atlassian-Token: no-check"" --location --request POST "
        strSend = strSend & strApiUri
        strSend = strSend & " --form ""file=@@sourcepath"""
        
        strAuth = EncodeBase64(Environ("username") + ":" + WINDOWS_PASSWORD)
        strSend = Replace(strSend, "@authorization", strAuth)
        
        strSend = Replace(strSend, "@sourcepath", strPath)
        
        'Execute in command line using windows script host
    
        wsh.Run strSend, 0, True
    
    Next dictKey
    
    MsgBox Prompt:="Successfully created JIRA tickets!"
    
endSubHandler:
    
'3.0 Tear down
    
    Set olApp = Nothing
    Set olItem = Nothing
    Set olItems = Nothing
    Set olTrgFldr = Nothing
    
    Set wsh = Nothing
    
    Set dict = Nothing
    
End Sub

Function GetEpicID(strSender As String) As String
'Description:____________________________________________________________________________________________________________________________________________________________________________________________________________________
'
'   - Returns Jira Epic ID based on user name
'

    Select Case strSender
            
        Case "Elisabeth Gloster", "Brooke Iritz", "Jorn-Caj Masters", "Anna Strnisha", "Morgan Coppola", "Veronica Addai", _
          "Nchang-Nwi Fon", "Sharise Owens", "Alexis Butler", "Kalliopi Tzantzara", "Olivia Molella"
        
            GetEpicID = "LUPBS-18"  'Contracts
        
        Case "Khaleda Akter", "Mandy Heinrich", "Dominique Hawkins"
        
            GetEpicID = "LUPBS-19"  'Master Data
            
        Case "Sarah Broes", "Nathan Skemp", "Rachel Nelson", "Marc Felder"
            
            GetEpicID = "LUPBS-464" 'ICIX
            
        Case "Colleen Elvin"
        
            GetEpicID = "LUPBS-497" 'Non-Food, Promo, Imports
        
        Case Else
        
            GetEpicID = "LUPBS-20"  'Buying
    
    End Select

End Function

Sub PostHTTPRequest(ByVal strApiUri As String, ByVal strSend As String, Optional ByRef strIssKey As String = "")
'Description:____________________________________________________________________________________________________________________________________________________________________________________________________________________
'
'   - Communicate provided JSON information to JIRA using an XML HTTP Request object
'
    
    Dim objXHR As New MSXML2.XMLHTTP60 'XML HTTP Request object
    Dim lngPos As Long
    
    With objXHR
        
        .Open "POST", strApiUri, False
        
        .setRequestHeader "Authorization", "Basic " & strAuth  'Encoded authorization string
        .setRequestHeader "X-Atlassian-Token", "no-check"
        .setRequestHeader "Content-Type", "application/json"
        .Send (strSend)

        Debug.Print .responseText
        Debug.Print .Status & " " & .statusText
        
        If .Status <> 200 And .Status <> 201 Then
            MsgBox Prompt:="Something went wrong. Communication with JIRA failed.", Buttons:=vbCritical
            End
        End If
        
        'Find JIRA key of newly-created issue
        
        lngPos = InStr(1, .responseText, "key") + 6
        strIssKey = Mid(.responseText, lngPos, InStr(lngPos, .responseText, ",") - 1 - lngPos)
        
    End With
    
    Set objXHR = Nothing

End Sub

Function ExtractNumber(StringToParse As String) As Double
    
    Dim Position As Integer
    Dim ParsedString As String, CharacterAtPosition As String
    
    'Loop through all characters in text string
    
    For Position = 1 To Len(StringToParse)
    
        CharacterAtPosition = Mid(StringToParse, Position, 1)
        
        'If the character is non-numeric and non-decimal, skip
        'Otherwise, add to output string
        
        If Not IsNumeric(CharacterAtPosition) Then
            If CharacterAtPosition <> "." Or InStr(1, ParsedString, ".") > 0 Then   'Only one decimal is allowed
                GoTo NextCharacter
            End If
        End If
        
        ParsedString = ParsedString & CharacterAtPosition
        
NextCharacter:

    Next Position
    
    If ParsedString <> "" Then ExtractNumber = ParsedString
    
End Function

Function CreateStrJSON(strSumm As String, strDesc As String, strEpicId As String, strPrior As String) As String
'Description:____________________________________________________________________________________________________________________________________________________________________________________________________________________
'
'   - Creates a JSON (Javascript Object Notation) string used to communicate information to JIRA
'

    Dim strTmp As String
    
    strTmp = strTmp & "{" & vbCr
    strTmp = strTmp & "   ""fields"":{" & vbCr
    strTmp = strTmp & "      ""project"":{" & vbCr
    strTmp = strTmp & "         ""id"":""@projectid""" & vbCr
    strTmp = strTmp & "      }," & vbCr
    strTmp = strTmp & "      ""summary"":""@summary""," & vbCr
    strTmp = strTmp & "      ""reporter"":{""name"":""@user""}," & vbCr
    strTmp = strTmp & "      ""assignee"":{""name"":""@user""}," & vbCr
    strTmp = strTmp & "      ""description"":""@description""," & vbCr
    strTmp = strTmp & "      ""priority"":{""name"":""@priority""}," & vbCr
    strTmp = strTmp & "      ""customfield_10000"":""@epicid""," & vbCr
    strTmp = strTmp & "      ""issuetype"":{" & vbCr
    strTmp = strTmp & "         ""id"":""@issuetype""" & vbCr
    strTmp = strTmp & "      }," & vbCr
    strTmp = strTmp & "      ""customfield_18310"":""@requestor""," & vbCr   'Requestor
    strTmp = strTmp & "      ""timetracking"":{" & vbCr
    strTmp = strTmp & "         ""originalEstimate"":""30m""" & vbCr
    strTmp = strTmp & "      }" & vbCr
    strTmp = strTmp & "   }" & vbCr
    strTmp = strTmp & "}"
    
    strTmp = Replace(strTmp, "@projectid", CStr(PROJECT_ID))
    strTmp = Replace(strTmp, "@user", Environ("username"))
    strTmp = Replace(strTmp, "@issuetype", CStr(ISSUE_TYPE_ID))
    strTmp = Replace(strTmp, "@summary", strSumm)
    strTmp = Replace(strTmp, "@description", strDesc)
    strTmp = Replace(strTmp, "@epicid", strEpicId)
    strTmp = Replace(strTmp, "@priority", strPrior)
    strTmp = Replace(strTmp, "@requestor", FirstLastName(Trim(Left(strSumm, InStr(1, strSumm, "-") - 1))))
    
    CreateStrJSON = strTmp

End Function

Public Function EncodeBase64(strTmp As String) As String

    Dim arrData() As Byte
    arrData = StrConv(strTmp, vbFromUnicode)
 
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement
 
    Set objXML = New MSXML2.DOMDocument60
    Set objNode = objXML.createElement("b64")
 
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.Text
 
    Set objNode = Nothing
    Set objXML = Nothing
    
End Function

Sub CreateBiweeklyTask()

    Dim strSumm As String, strDesc As String, strJSON As String
    
    strSumm = "Biweekly Tasks"
    
    strDesc = strDesc & "* Approved Archive Backup\n"
    strDesc = strDesc & "* Approved Archive Audit\n"
    strDesc = strDesc & "\n"
    strDesc = strDesc & "* Missing Items\n"
    strDesc = strDesc & "\n"
    strDesc = strDesc & "* New SAPs\n"
    strDesc = strDesc & "* Row Count\n"
    strDesc = strDesc & "* New Suppliers\n"
    strDesc = strDesc & "* EKS Table\n"
    strDesc = strDesc & "\n"
    strDesc = strDesc & "* Neg Log Changes\n"
    strDesc = strDesc & "\n"
    
    strJSON = CreateStrJSON(strSumm, strDesc, "LUPBS-16", "Highest")
    
    Call PostHTTPRequest(API_URI, strJSON)
    
End Sub

Sub CreateReviewPreviewTask()

    Dim strSumm As String, strDesc As String, strJSON As String
    
    strSumm = "Review Preview"
    
    strDesc = "Update power BI workflow summary"

    strJSON = CreateStrJSON(strSumm, strDesc, "LUPBS-16", "Highest")
    
    Call PostHTTPRequest(API_URI, strJSON)
    
End Sub

Sub AutomateJIRA()

    Dim ws As Worksheet
    Dim x As Integer
    Dim strTime As String, strProc As String
    
    Set ws = ThisWorkbook.Worksheets("Task Schedule")
    
    For x = 0 To 51
    
        strProc = "CreateBiweeklyTask"
        strTime = CStr(CDate(CDbl(CDate("9/9/2021 09:30:00 AM")) + 7 * x))
        
        Application.OnTime EarliestTime:=strTime, Procedure:=strProc
        
        With ws
            .Range("A" & .Cells(.Rows.Count, 1).End(xlUp).Row + 1).Value = strTime
            .Range("B" & .Cells(.Rows.Count, 2).End(xlUp).Row + 1).Value = strProc
        End With
    
    Next x
    
    Set ws = Nothing

End Sub

Function FirstLastName(NameString As String) As String
    If InStr(1, NameString, ",") = 0 Then FirstLastName = NameString: Exit Function
    Dim LastFirst As Variant: LastFirst = Split(NameString, ",")
    FirstLastName = Trim(LastFirst(1)) & " " & Trim(LastFirst(0))
End Function

