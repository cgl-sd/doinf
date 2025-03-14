Attribute VB_Name = "FetchDOIInfo"
Sub DoiInfo()
    ' ��������
    Dim doi As String          ' �洢DOI��
    Dim url As String          ' �洢API����URL
    Dim http As Object         ' HTTP�������
    Dim response As String     ' �洢API���ص�JSON��Ӧ
    Dim json As Object         ' �������JSON����
    Dim title As String        ' ���ױ���
    Dim pubDate As String      ' ��������
    Dim journal As String      ' �ڿ���
    Dim firstAuthorGiven As String, firstAuthorFamily As String, fullName As String '��һ����ȫ��
    Dim dateParts As Collection
    Dim year As Integer, month As Integer, day As Integer
    Dim formattedDate As String
    
    
    ' ������
    On Error GoTo ErrorHandler
    
    ' ��ȡ��ǰѡ�е�Ԫ���DOI��
    doi = ActiveCell.Value
    If doi = "" Then Exit Sub  ' �����Ԫ��Ϊ�����˳�
    
    ' ����CrossRef API����URL
    url = "https://api.crossref.org/works/" & doi
    
    ' ����HTTP GET����
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.send
    
    ' ���HTTP״̬��
    If http.Status = 200 Then
        response = http.responseText
        Set json = JsonConverter.ParseJson(response)  ' ����JSON
        
        
        ' ��һ�У������������ݣ�publish print���ȣ�
        With json("message")
            Set dateParts = .Item("published-print")("date-parts")(1) ' ��һ����������
        End With
        year = dateParts(1)   ' ��
        month = dateParts(2)  ' ��
        ' �������ȱʧ�յ����
        If dateParts.Count >= 3 Then
            day = dateParts(3)
        Else
            day = 1 ' Ĭ����Ϊ���µ�һ��
        End If
        formattedDate = Format(DateSerial(year, month, day), "yyyy-mm-dd")
        ActiveCell.offset(0, 1).Value = formattedDate
        
        '�ڶ��У���õ�һ���߲�ƴ��ȫ��
        firstAuthorGiven = json("message")("author")(1)("given")
        firstAuthorFamily = json("message")("author")(1)("family")
        fullName = firstAuthorGiven & " " & firstAuthorFamily
        ActiveCell.offset(0, 2).Value = fullName
        
        '���������У��ڿ����ͱ���
        journal = json("message")("short-container-title")(1)
        ActiveCell.offset(0, 3).Value = journal '
        title = json("message")("title")(1)
        ActiveCell.offset(0, 4).Value = title
        
    Else
        ' ��ʾHTTP����
        ActiveCell.offset(0, 5).Value = "Error: " & http.Status
    End If
    
    ' ������Դ
    Set http = Nothing
    Application.Wait Now + TimeValue("00:00:01")  ' �ӳ�1���API����
    Exit Sub

ErrorHandler:
    MsgBox "����ԭ��" & Err.Description & vbCrLf & "��飺1. DOI�Ƿ���ȷ 2. �����Ƿ���ͨ", , "DOI�������"
End Sub
