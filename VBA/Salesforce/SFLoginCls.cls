VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SFLoginCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Const LOGIN_URL = "https://login.salesforce.com/services/Soap/u/47.0"
    Const XML_PATH = "C:\Users\tusha\Documents\projects\MyGithub\tusharthomas\VBA\Salesforce\logindata.xml"
    
    Private pResponse As String

    Public Function Login(Username As String, Password As String, Secret As String) As String
    
        'Submits login request and returns session ID
        
        On Error Resume Next

            Login = _
                ProcessResponse( _
                    SubmitPOSTRequest( _
                        GetXMLLoginInfo(Username, Password, Secret) _
                    ) _
                )
                
            If Err.Number > 0 Then Debug.Print "Warning! Something went wrong with the login. Could not get open session with Salesforce. See response below:" & vbCrLf & "[" & pResponse & "]"

        On Error GoTo 0

    End Function

    Private Function GetXMLLoginInfo(Username As String, Password As String, SecurityToken As String) As String

        Dim XMLDoc As MSXML2.DOMDocument60:     Set XMLDoc = New MSXML2.DOMDocument60
        
        Dim strTmp As String
        Dim Values As Variant

        Values = _
            Array( _
                "@username:" & Username, _
                "@password:" & Password, _
                "@secret:" & SecurityToken)

        strTmp = FileOps.ReadContentsFromFile(XML_PATH, Values)
        XMLDoc.LoadXML strTmp

        GetXMLLoginInfo = XMLDoc.XML

    End Function

    Private Function SubmitPOSTRequest(RequestData As String) As String

        'Submits request to Salesforce and returns response text

        Dim Request As MSXML2.XMLHTTP60:    Set Request = New MSXML2.XMLHTTP60

        With Request

            .Open "POST", LOGIN_URL, False
            .setRequestHeader "Content-Type", "text/xml"
            .setRequestHeader "SOAPAction", "login"
            .send RequestData

            pResponse = .responseText: SubmitPOSTRequest = pResponse
            
            If .Status <> "200" Then Err.Raise 513

        End With

    End Function

    Private Function ProcessResponse(Response As String) As String

        Dim XMLDoc As MSXML2.DOMDocument60:     Set XMLDoc = New MSXML2.DOMDocument60

        XMLDoc.LoadXML Response
        XMLDoc.SetProperty "SelectionNamespaces", "xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:urn=""urn:partner.soap.sforce.com"""
        ProcessResponse = XMLDoc.SelectSingleNode("//urn:result/urn:sessionId").Text

    End Function
