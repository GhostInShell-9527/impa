# impa, excel Macro
check IMPA code in webpage


Sub OpenWebPage()
    Dim SelectedCell As Range
    Dim BasePath As String
    Dim Code As String
    Dim FullPath As String
    
    ' 定义文件路径的基础路径
    BasePath = "D:\BRUCE\BRUCE\k\1\1\www.space-marine.co.kr\pricelist\"
    
    ' 遍历选定的单元格
    For Each SelectedCell In Selection
        ' 检查单元格是否为数字，且长度为6位
        If IsNumeric(SelectedCell.Value) And Len(SelectedCell.Value) = 6 Then
            ' 获取代码
            Code = SelectedCell.Value
            
            ' 拼接完整路径，假设 category 是代码的前两位
            FullPath = BasePath & "information.cfm-code=" & Code & "&category=" & Left(Code, 2) & ".htm"
            
            ' 使用 Shell 调用默认浏览器打开网页
            ThisWorkbook.FollowHyperlink FullPath

        Else
            MsgBox "选中的单元格包含无效代码: " & SelectedCell.Value, vbExclamation, "错误"
        End If
    Next SelectedCell
End Sub
