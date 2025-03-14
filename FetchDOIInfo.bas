Attribute VB_Name = "FetchDOIInfo"
Sub DoiInfo()
    ' 声明变量
    Dim doi As String          ' 存储DOI号
    Dim url As String          ' 存储API请求URL
    Dim http As Object         ' HTTP请求对象
    Dim response As String     ' 存储API返回的JSON响应
    Dim json As Object         ' 解析后的JSON对象
    Dim title As String        ' 文献标题
    Dim pubDate As String      ' 出版日期
    Dim journal As String      ' 期刊名
    Dim firstAuthorGiven As String, firstAuthorFamily As String, fullName As String '第一作者全名
    Dim dateParts As Collection
    Dim year As Integer, month As Integer, day As Integer
    Dim formattedDate As String
    
    
    ' 错误处理
    On Error GoTo ErrorHandler
    
    ' 获取当前选中单元格的DOI号
    doi = ActiveCell.Value
    If doi = "" Then Exit Sub  ' 如果单元格为空则退出
    
    ' 构建CrossRef API请求URL
    url = "https://api.crossref.org/works/" & doi
    
    ' 发送HTTP GET请求
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.send
    
    ' 检查HTTP状态码
    If http.Status = 200 Then
        response = http.responseText
        Set json = JsonConverter.ParseJson(response)  ' 解析JSON
        
        
        ' 第一列，导航日期数据（publish print优先）
        With json("message")
            Set dateParts = .Item("published-print")("date-parts")(1) ' 第一组日期数据
        End With
        year = dateParts(1)   ' 年
        month = dateParts(2)  ' 月
        ' 处理可能缺失日的情况
        If dateParts.Count >= 3 Then
            day = dateParts(3)
        Else
            day = 1 ' 默认设为当月第一天
        End If
        formattedDate = Format(DateSerial(year, month, day), "yyyy-mm-dd")
        ActiveCell.offset(0, 1).Value = formattedDate
        
        '第二列，获得第一作者并拼接全名
        firstAuthorGiven = json("message")("author")(1)("given")
        firstAuthorFamily = json("message")("author")(1)("family")
        fullName = firstAuthorGiven & " " & firstAuthorFamily
        ActiveCell.offset(0, 2).Value = fullName
        
        '第三、四列，期刊名和标题
        journal = json("message")("short-container-title")(1)
        ActiveCell.offset(0, 3).Value = journal '
        title = json("message")("title")(1)
        ActiveCell.offset(0, 4).Value = title
        
    Else
        ' 显示HTTP错误
        ActiveCell.offset(0, 5).Value = "Error: " & http.Status
    End If
    
    ' 清理资源
    Set http = Nothing
    Application.Wait Now + TimeValue("00:00:01")  ' 延迟1秒防API限流
    Exit Sub

ErrorHandler:
    MsgBox "出错原因：" & Err.Description & vbCrLf & "检查：1. DOI是否正确 2. 网络是否连通", , "DOI导入错误"
End Sub
