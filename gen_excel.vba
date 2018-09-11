Private Sub save_xlsx(ByVal fileFullName As String)
    Application.DisplayAlerts = False
    '打开副本
    Dim wb
    Set wb = Workbooks.Open(fileFullName)
    wb.Activate
    Dim replaceFile    As String
    replaceFile = Replace(fileFullName, ".xlsm", ".xlsx")
    ActiveWorkbook.SaveAs fileName:=replaceFile, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    '退出保存
    wb.Close SaveChanges:=False
    If Dir(fileFullName, 16) <> Empty Then
		On Error Resume Next
        Kill fileFullName
    End If
End Sub

Sub a1()
    Application.ScreenUpdating = False

    '变量声明
    Dim fileFullName    As String    '工作簿全路径的文件名
    Dim FilePath        As String    '工作簿路径
    Dim fileName        As String    '工作簿的文件名
    '文件名
    fileName = "愛知製油所"
    '获取路径
    FilePath = ThisWorkbook.Path & "\gen\"
    '创建目录
    If Dir(FilePath, vbDirectory) = "" Then
        MkDir (FilePath)
    End If
    '获取全路径文件名
    fileFullName = FilePath & fileName & ".xlsm"
    If Dir(fileFullName, 16) <> Empty Then
		On Error Resume Next
        Kill fileFullName
    End If
    '创建副本
    ThisWorkbook.SaveCopyAs fileFullName
    
    '打开副本
    Dim wb
    Set wb = Workbooks.Open(fileFullName)
    wb.Activate
    '删除不需要的行
    Dim Index As Long
    Dim context As String
    For Index = Cells(Rows.Count, "B").End(xlUp).Row To 9 Step -1
        context = ActiveSheet.Cells(Index, "AJ")
        Debug.Print context
        If VBA.Trim("愛知製油所") = VBA.Trim(context) = False Then
            Debug.Print Index
            Rows(Index).Delete Shift:=xlShiftUp
        End If
    Next
    ActiveSheet.Cells(3, "B").Value = "愛知製油所"
    '删除按钮
    ActiveSheet.Shapes.Range(Array("Button 10")).Select
    Selection.Delete
    '隐藏列
    Columns("AG:AR").Select
    Range("AG9").Activate
    Selection.EntireColumn.Hidden = True
    '加密码
    Columns("AG:AR").Select
    Range("AG9").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=6303

    '退出保存
    wb.Close SaveChanges:=True
    
    '换格式
    Call save_xlsx(fileFullName)
End Sub

Sub a2()
    Application.ScreenUpdating = False

    '变量声明
    Dim fileFullName    As String    '工作簿全路径的文件名
    Dim FilePath        As String    '工作簿路径
    Dim fileName        As String    '工作簿的文件名
    '文件名
    fileName = "徳山事業所"
    '获取路径
    FilePath = ThisWorkbook.Path & "\gen\"
    '创建目录
    If Dir(FilePath, vbDirectory) = "" Then
        MkDir (FilePath)
    End If
    '获取全路径文件名
    fileFullName = FilePath & fileName & ".xlsm"
    If Dir(fileFullName, 16) <> Empty Then
		On Error Resume Next
        Kill fileFullName
    End If
    '创建副本
    ThisWorkbook.SaveCopyAs fileFullName
    
    '打开副本
    Dim wb
    Set wb = Workbooks.Open(fileFullName)
    wb.Activate

    Dim Index As Long
    Dim context As String
    For Index = Cells(Rows.Count, "B").End(xlUp).Row To 9 Step -1
        context = ActiveSheet.Cells(Index, "AJ")
        Debug.Print context
        If VBA.Trim("徳山事業所") = VBA.Trim(context) = False Then
            Debug.Print Index
            Rows(Index).Delete Shift:=xlShiftUp
        End If
    Next
    ActiveSheet.Cells(3, "B").Value = "徳山事業所"
    '删除按钮
    ActiveSheet.Shapes.Range(Array("Button 10")).Select
    Selection.Delete
    '隐藏列
    Columns("AG:AR").Select
    Range("AG9").Activate
    Selection.EntireColumn.Hidden = True
    '加密码
    Columns("AG:AR").Select
    Range("AG9").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=6303
    '退出保存
    wb.Close SaveChanges:=True
    '换格式
    Call save_xlsx(fileFullName)
End Sub


Sub a3()
    Application.ScreenUpdating = False

    '变量声明
    Dim fileFullName    As String    '工作簿全路径的文件名
    Dim FilePath        As String    '工作簿路径
    Dim fileName        As String    '工作簿的文件名
    '文件名
    fileName = "北海道製油所"
    '获取路径
    FilePath = ThisWorkbook.Path & "\gen\"
    '创建目录
    If Dir(FilePath, vbDirectory) = "" Then
        MkDir (FilePath)
    End If
    '获取全路径文件名
    fileFullName = FilePath & fileName & ".xlsm"
    If Dir(fileFullName, 16) <> Empty Then
		On Error Resume Next
        Kill fileFullName
    End If
    '创建副本
    ThisWorkbook.SaveCopyAs fileFullName
    
    '打开副本
    Dim wb
    Set wb = Workbooks.Open(fileFullName)
    wb.Activate

    Dim Index As Long
    Dim context As String
    For Index = Cells(Rows.Count, "B").End(xlUp).Row To 9 Step -1
        context = ActiveSheet.Cells(Index, "AJ")
        Debug.Print context
        If VBA.Trim("北海道製油所") = VBA.Trim(context) = False Then
            Debug.Print Index
            Rows(Index).Delete Shift:=xlShiftUp
        End If
    Next
    ActiveSheet.Cells(3, "B").Value = "北海道製油所"
    '删除按钮
    ActiveSheet.Shapes.Range(Array("Button 10")).Select
    Selection.Delete
    '隐藏列
    Columns("AG:AR").Select
    Range("AG9").Activate
    Selection.EntireColumn.Hidden = True
    '加密码
    Columns("AG:AR").Select
    Range("AG9").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=6303
    '退出保存
    wb.Close SaveChanges:=True
    '换格式
    Call save_xlsx(fileFullName)
End Sub

Sub a4()
    Application.ScreenUpdating = False

    '变量声明
    Dim fileFullName    As String    '工作簿全路径的文件名
    Dim FilePath        As String    '工作簿路径
    Dim fileName        As String    '工作簿的文件名
    '文件名
    fileName = "千葉事業所"
    '获取路径
    FilePath = ThisWorkbook.Path & "\gen\"
    '创建目录
    If Dir(FilePath, vbDirectory) = "" Then
        MkDir (FilePath)
    End If
    '获取全路径文件名
    fileFullName = FilePath & fileName & ".xlsm"
    If Dir(fileFullName, 16) <> Empty Then
		On Error Resume Next
        Kill fileFullName
    End If
    '创建副本
    ThisWorkbook.SaveCopyAs fileFullName
    
    '打开副本
    Dim wb
    Set wb = Workbooks.Open(fileFullName)
    wb.Activate

    Dim Index As Long
    Dim context As String
    For Index = Cells(Rows.Count, "B").End(xlUp).Row To 9 Step -1
        context = ActiveSheet.Cells(Index, "AJ")
        Debug.Print context
        If Not VBA.Trim("千葉事業所（化学）") = VBA.Trim(context) And Not VBA.Trim("千葉事業所（石油）") = VBA.Trim(context) Then
            Debug.Print Index
            Rows(Index).Delete Shift:=xlShiftUp
        End If
    Next
    ActiveSheet.Cells(3, "B").Value = "千葉事業所"
    '删除按钮
    ActiveSheet.Shapes.Range(Array("Button 10")).Select
    Selection.Delete
    '隐藏列
    Columns("AG:AR").Select
    Range("AG9").Activate
    Selection.EntireColumn.Hidden = True
    '加密码
    Columns("AG:AR").Select
    Range("AG9").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=6303
    '退出保存
    wb.Close SaveChanges:=True
    '换格式
    Call save_xlsx(fileFullName)
End Sub

Sub b1()
    Application.ScreenUpdating = False

    '变量声明
    Dim fileFullName    As String    '工作簿全路径的文件名
    Dim FilePath        As String    '工作簿路径
    Dim fileName        As String    '工作簿的文件名
    '文件名
    fileName = "未収金 諸口"
    '获取路径
    FilePath = ThisWorkbook.Path & "\gen\"
    '创建目录
    If Dir(FilePath, vbDirectory) = "" Then
        MkDir (FilePath)
    End If
    '获取全路径文件名
    fileFullName = FilePath & fileName & ".xlsm"
    If Dir(fileFullName, 16) <> Empty Then
		On Error Resume Next
        Kill fileFullName
    End If
    '创建副本
    ThisWorkbook.SaveCopyAs fileFullName
    
    '打开副本
    Dim wb
    Set wb = Workbooks.Open(fileFullName)
    wb.Activate

    Dim Index As Long
    Dim context1 As String
    Dim context2 As String
    For Index = Cells(Rows.Count, "B").End(xlUp).Row To 9 Step -1
        context1 = ActiveSheet.Cells(Index, "C")
        context2 = ActiveSheet.Cells(Index, "AJ")
        If (VBA.Trim("未収金 諸口") = VBA.Trim(context1)) And (VBA.Trim("本社") = VBA.Trim(context2)) Then
            Debug.Print Index
        Else
            Debug.Print Index
            Rows(Index).Delete Shift:=xlShiftUp
        End If
    Next
    ActiveSheet.Cells(Index, "B").Value = "本社"
    '删除按钮
    ActiveSheet.Shapes.Range(Array("Button 10")).Select
    Selection.Delete
    '隐藏列
    Columns("AG:AR").Select
    Range("AG9").Activate
    Selection.EntireColumn.Hidden = True
    '加密码
    Columns("AG:AR").Select
    Range("AG9").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=6303
    '退出保存
    wb.Close SaveChanges:=True
    '换格式
    Call save_xlsx(fileFullName)
End Sub

Private Sub GenDep(ByVal DepName As String)
    Application.ScreenUpdating = False

    '变量声明
    Dim fileFullName    As String    '工作簿全路径的文件名
    Dim FilePath        As String    '工作簿路径
    Dim fileName        As String    '工作簿的文件名
    '文件名
    fileName = DepName
    '获取路径
    FilePath = ThisWorkbook.Path & "\gen\"
    '创建目录
    If Dir(FilePath, vbDirectory) = "" Then
        MkDir (FilePath)
    End If
    '获取全路径文件名
    fileFullName = FilePath & fileName & ".xlsm"
    If Dir(fileFullName, 16) <> Empty Then
		On Error Resume Next
        Kill fileFullName
    End If
    '创建副本
    ThisWorkbook.SaveCopyAs fileFullName
    
    '打开副本
    Dim wb
    Set wb = Workbooks.Open(fileFullName)
    wb.Activate

    Dim Index As Long
    Dim context1 As String
    Dim context2 As String
    For Index = Cells(Rows.Count, "B").End(xlUp).Row To 9 Step -1
        context1 = ActiveSheet.Cells(Index, "AJ")
        context2 = ActiveSheet.Cells(Index, "AK")
        If (VBA.Trim("愛知製油所") = VBA.Trim(context1)) Or _
        (VBA.Trim("徳山事業所") = VBA.Trim(context1)) Or _
        (VBA.Trim("北海道製油所") = VBA.Trim(context1)) Or _
        (VBA.Trim("千葉事業所（石油）") = VBA.Trim(context1)) Or _
        (VBA.Trim("千葉事業所（化学）") = VBA.Trim(context1)) Or _
        (VBA.Trim("") = VBA.Trim(context1)) Then
            Debug.Print Index
            Rows(Index).Delete Shift:=xlShiftUp
        Else
            If Not VBA.Trim(DepName) = VBA.Trim(context2) Then
                Rows(Index).Delete Shift:=xlShiftUp
            End If
        End If
    Next
    ActiveSheet.Cells(Index, "B").Value = "本社"
    '删除按钮
    ActiveSheet.Shapes.Range(Array("Button 10")).Select
    Selection.Delete
    '隐藏列
    Columns("AG:AR").Select
    Range("AG9").Activate
    Selection.EntireColumn.Hidden = True
    '加密码
    Columns("AG:AR").Select
    Range("AG9").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=6303
    '退出保存
    wb.Close SaveChanges:=True
    '换格式
    Call save_xlsx(fileFullName)
End Sub
Private Sub GenDep_c1(ByVal DepName As String)
    Application.ScreenUpdating = False

    '变量声明
    Dim fileFullName    As String    '工作簿全路径的文件名
    Dim FilePath        As String    '工作簿路径
    Dim fileName        As String    '工作簿的文件名
    '文件名
    fileName = DepName
    '获取路径
    FilePath = ThisWorkbook.Path & "\gen\"
    '创建目录
    If Dir(FilePath, vbDirectory) = "" Then
        MkDir (FilePath)
    End If
    '获取全路径文件名
    fileFullName = FilePath & fileName & ".xlsm"
    If Dir(fileFullName, 16) <> Empty Then
		On Error Resume Next
        Kill fileFullName
    End If
    '创建副本
    ThisWorkbook.SaveCopyAs fileFullName
    
    '打开副本
    Dim wb
    Set wb = Workbooks.Open(fileFullName)
    wb.Activate

    Dim Index As Long
    Dim context1 As String
    Dim context2 As String
    For Index = Cells(Rows.Count, "B").End(xlUp).Row To 9 Step -1
        context1 = ActiveSheet.Cells(Index, "AJ")
        context2 = ActiveSheet.Cells(Index, "AK")
        context3 = ActiveSheet.Cells(Index, "AL")
        If (VBA.Trim("愛知製油所") = VBA.Trim(context1)) Or _
        (VBA.Trim("徳山事業所") = VBA.Trim(context1)) Or _
        (VBA.Trim("北海道製油所") = VBA.Trim(context1)) Or _
        (VBA.Trim("千葉事業所（石油）") = VBA.Trim(context1)) Or _
        (VBA.Trim("千葉事業所（化学）") = VBA.Trim(context1)) Or _
        (VBA.Trim("") = VBA.Trim(context1)) Then
            Debug.Print Index
            Rows(Index).Delete Shift:=xlShiftUp
        Else
            If (Not VBA.Trim("潤滑油一部") = VBA.Trim(context2)) And (Not VBA.Trim("潤滑油二部") = VBA.Trim(context2)) Then
                Rows(Index).Delete Shift:=xlShiftUp
            Else
                If VBA.Trim("営業研究所") = VBA.Trim(context3) Then
                    Rows(Index).Delete Shift:=xlShiftUp
                End If
            End If
        End If
    Next
    ActiveSheet.Cells(Index, "B").Value = "本社"
    '删除按钮
    ActiveSheet.Shapes.Range(Array("Button 10")).Select
    Selection.Delete
    '隐藏列
    Columns("AG:AR").Select
    Range("AG9").Activate
    Selection.EntireColumn.Hidden = True
    '加密码
    Columns("AG:AR").Select
    Range("AG9").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=6303
    '退出保存
    wb.Close SaveChanges:=True
    '换格式
    Call save_xlsx(fileFullName)
End Sub
Private Sub GenDep_c2(ByVal DepName As String)
    Application.ScreenUpdating = False

    '变量声明
    Dim fileFullName    As String    '工作簿全路径的文件名
    Dim FilePath        As String    '工作簿路径
    Dim fileName        As String    '工作簿的文件名
    '文件名
    fileName = DepName
    '获取路径
    FilePath = ThisWorkbook.Path & "\gen\"
    '创建目录
    If Dir(FilePath, vbDirectory) = "" Then
        MkDir (FilePath)
    End If
    '获取全路径文件名
    fileFullName = FilePath & fileName & ".xlsm"
    If Dir(fileFullName, 16) <> Empty Then
		On Error Resume Next
        Kill fileFullName
    End If
    '创建副本
    ThisWorkbook.SaveCopyAs fileFullName
    
    '打开副本
    Dim wb
    Set wb = Workbooks.Open(fileFullName)
    wb.Activate

    Dim Index As Long
    Dim context1 As String
    Dim context2 As String
    For Index = Cells(Rows.Count, "B").End(xlUp).Row To 9 Step -1
        context1 = ActiveSheet.Cells(Index, "AJ")
        context2 = ActiveSheet.Cells(Index, "AK")
        context3 = ActiveSheet.Cells(Index, "AL")
        If (VBA.Trim("愛知製油所") = VBA.Trim(context1)) Or _
        (VBA.Trim("徳山事業所") = VBA.Trim(context1)) Or _
        (VBA.Trim("北海道製油所") = VBA.Trim(context1)) Or _
        (VBA.Trim("千葉事業所（石油）") = VBA.Trim(context1)) Or _
        (VBA.Trim("千葉事業所（化学）") = VBA.Trim(context1)) Or _
        (VBA.Trim("") = VBA.Trim(context1)) Then
            Debug.Print Index
            Rows(Index).Delete Shift:=xlShiftUp
        Else
            If (Not VBA.Trim("潤滑油一部") = VBA.Trim(context2)) And (Not VBA.Trim("潤滑油二部") = VBA.Trim(context2)) Then
                Rows(Index).Delete Shift:=xlShiftUp
            Else
                If Not VBA.Trim("営業研究所") = VBA.Trim(context3) Then
                    Rows(Index).Delete Shift:=xlShiftUp
                End If
            End If
        End If
    Next
    ActiveSheet.Cells(Index, "B").Value = "本社"
    '删除按钮
    ActiveSheet.Shapes.Range(Array("Button 10")).Select
    Selection.Delete
    '隐藏列
    Columns("AG:AR").Select
    Range("AG9").Activate
    Selection.EntireColumn.Hidden = True
    '加密码
    Columns("AG:AR").Select
    Range("AG9").Activate
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=6303
    '退出保存
    wb.Close SaveChanges:=True
    '换格式
    Call save_xlsx(fileFullName)
End Sub

Sub all_dep()
    Application.ScreenUpdating = False

    Dim dc As Object
    Set dc = CreateObject("Scripting.dictionary")

    Dim Index As Long
    Dim context As String
    For Index = Cells(Rows.Count, "B").End(xlUp).Row To 9 Step -1
        context = ActiveSheet.Cells(Index, "AK")
        If Not dc.exists(context) Then
            dc.Add context, ""
        End If
    Next
    
    Dim ar
    ar = dc.keys
    For i = 0 To UBound(ar)
    
            If (Not VBA.Trim("潤滑油一部") = VBA.Trim(ar(i))) And _
            (Not VBA.Trim("潤滑油二部") = VBA.Trim(ar(i))) And _
            (Not VBA.Trim("愛知製油所") = VBA.Trim(ar(i))) And _
            (Not VBA.Trim("徳山事業所") = VBA.Trim(ar(i))) And _
            (Not VBA.Trim("北海道製油所") = VBA.Trim(ar(i))) And _
            (Not VBA.Trim("千葉事業所（石油）") = VBA.Trim(ar(i))) And _
            (Not VBA.Trim("千葉事業所（化学）") = VBA.Trim(ar(i))) And _
            (Not VBA.Trim("") = VBA.Trim(ar(i))) Then
                Call GenDep(ar(i))
            End If
    Next
    Call GenDep_c1("潤滑油部")
    Call GenDep_c2("営業研究所")
End Sub
Sub all()
    Application.ScreenUpdating = False
    Call a1
    Call a2
    Call a3
    Call a4
    Call b1
    Call all_dep
    MsgBox "all finishd"
End Sub


