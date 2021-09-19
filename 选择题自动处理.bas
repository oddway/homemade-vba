Attribute VB_Name = "选择题自动处理"
Sub 选择题选项重排()
    st = VBA.Timer

    '    Application.ScreenUpdating = False '关闭屏幕更新
    '    On Error Resume Next
    Dim s, x, Y, n, m, Mp, t As Single
    '    Dim aRange, bRange, cRange, dRange As Range
    Dim bSelect, cSelect, dSelect As Selection
    Dim H_Select, XX, TmpChr, AnserChr As String

    XX = "ABCD" '选择题选项
    '        XX = InputBox(Prompt:="请输入选择题选项的字母。", Default:="ABCD")

    
    s = Len(XX) '取选项字符串长度
    ActiveDocument.ActiveWindow.View.ShowAll = True '显示所有编辑标记

    For i = 1 To ActiveDocument.Paragraphs.count

        ActiveDocument.Range(ActiveDocument.Paragraphs(i).Range.Start, ActiveDocument.Paragraphs(i).Range.End).Select    '选定本段
        
        If InStrRev(ActiveDocument.Paragraphs(i).Range.text, "【　　】") Then
    '
    '            Selection.Find.ClearFormatting
    '            Selection.Find.MatchWildcards = True


            For m = 1 To Int(s / 2)

                ActiveDocument.Range(ActiveDocument.Paragraphs(i).Range.Start, ActiveDocument.Paragraphs(i).Range.End).Select    '选定本段


        Selection.Find.ClearFormatting
        With Selection.Find
            '        .Text = "[A-D]{1,2}"
            .MatchWildcards = True
            .Font.ColorIndex = wdRed
            .Font.Hidden = True

        End With
        Selection.Find.Execute findtext:="[" & XX & "]{1,2}"
        
        
    '                H_Select = Selection    '定义正确答案对象
                
                
                AnserChr = Selection    '正确答案字母
                
                
                
                TmpChr = Selection  '临时正确答案字母
    '            Selection.Find.Font.Hidden = False      '不查找隐藏内容
                    Selection.Find.ClearFormatting

                If Len(AnserChr) = 2 Then Exit For '如果给定答案是多字母，则跳过。
                
                Randomize (Timer)

                x = Int((4 * Rnd()) + 1) '生成1-4的随机数

                ActiveDocument.Range(ActiveDocument.Paragraphs(i).Range.Start, ActiveDocument.Paragraphs(i).Range.End).Select    '选定本段

                Selection.Find.Execute findtext:=Mid(XX, x, 1) & "[．.、]"  'mid函数：返回文本字符串中从指定位置开始的特定数目的字符，每次取一个数字

                Selection.MoveRight Unit:=wdCharacter, count:=1     '光标右移

                Ts = Selection.Start    '取光标所在为位置为选定起点TempStart

                Selection.Find.Execute findtext:="[^9^11^13]"

                Selection.MoveLeft Unit:=wdCharacter, count:=1     '光标移动

                Te = Selection.Start    '取光标所在为位置为选定终点TempEnd

                Selection.Start = Ts
                Selection.End = Te

                If Ts = Te Then GoTo k

                Selection.Cut    '第1次剪切内容

                Y = Int((4 * Rnd()) + 1) '生成1-4的随机数

                If Y = x Then Y = Int((4 * Rnd()) + 1)            '如果y=x，再生成一个1-4的随机数

                ActiveDocument.Range(ActiveDocument.Paragraphs(i).Range.Start, ActiveDocument.Paragraphs(i).Range.End).Select    '选定本段

                Selection.Find.Execute findtext:=Mid(XX, Y, 1) & "[．.、]"  'mid函数：返回ABCD中从指定位置开始的特定数目的字符，每次取1个数字

                If AnserChr = Mid(XX, x, 1) Then TmpChr = Mid(XX, Y, 1) '如果第一次剪切的答案正确，把答案字母临时更换这第二次随机的字母

                If AnserChr = Mid(XX, Y, 1) Then TmpChr = Mid(XX, x, 1) '如果第二次随机的答案正确，就把答案字母临时更换为第一次剪切项的字母



                Selection.MoveRight Unit:=wdCharacter, count:=1     '光标右移

                Selection.Paste          '粘贴第1次剪切内容

                Ts = Selection.Start    '取光标所在为位置为选定起点TempStart

                Selection.Find.Execute findtext:="[^9^11^13]"

                Selection.MoveLeft Unit:=wdCharacter, count:=1     '光标移动

                Te = Selection.Start    '取光标所在为位置为选定终点TempEnd

                Selection.Start = Ts
                Selection.End = Te

                If Ts = Te Then GoTo k

                Selection.Cut        '第2次剪切

                ActiveDocument.Range(ActiveDocument.Paragraphs(i).Range.Start, ActiveDocument.Paragraphs(i).Range.End).Select    '选定本段

                Selection.Find.Execute findtext:=Mid(XX, x, 1) & "[．.、]"  '本段第一次剪切的选项选定

                'Selection.Find.Execute findtext:="[A-D][．.、]"   '在已经选定部分查找并选定

                Selection.MoveRight Unit:=wdCharacter, count:=1     '光标右移到第一次剪切的选项后
                Selection.Paste             '粘贴第2次剪切内容

                
                If AnserChr = TmpChr Then GoTo k     '如果没有改变，直接跳过
                '把隐匿的答案字母更换
                ActiveDocument.Range(ActiveDocument.Paragraphs(i).Range.Start, ActiveDocument.Paragraphs(i).Range.End).Select    '选定本段

                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                With Selection.Find
                    .Font.Hidden = True          '只查找隐藏内容
                    .Font.Color = wdColorRed
                    
                    

                    .text = AnserChr
                    .Replacement.text = TmpChr

                    .MatchCase = False

                End With
                Selection.Find.Execute Replace:=wdReplaceAll

                Selection.Find.ClearFormatting

    '                Selection.Find.Font.Hidden = False      '不查找隐藏内容

    '                ActiveDocument.Range(ActiveDocument.Paragraphs(i).Range.Start, ActiveDocument.Paragraphs(i).Range.End).Select    '选定本段

k:
            Next m
            n = n + 1
        Else
        End If
    Next i

    '    If n Then
    '        MsgBox "共有" & n & "道选择题选项调整。"
    '    Else
    '        MsgBox "没有选择题选项调整。"
    '
    '    End If

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting

    With Selection.Find
        .text = " "
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        '        .Format = True
        '        .MatchCase = False
        '        .MatchWholeWord = False
        '        .MatchByte = False
        '        .MatchWildcards = False
        '        .MatchSoundsLike = False
        '        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    '    Application.ScreenUpdating = True '屏幕更新



    MsgBox "有 " & n & " 道选择题选项调整，用时 " & VBA.Format(Timer - st, "0.00") & " 秒"

End Sub


Sub 选择题选项重排1()


    '以下数字变量的单位均为“磅”
    'Application.ScreenUpdating = False '关闭屏幕更新

    '    On Error Resume Next


    Dim s, x, Y, n, m, Mp, t As Single
    Dim aRange, bRange, cRange, dRange As Range
    Dim aSelect, bSelect, cSelect, dSelect As Selection
    Dim XX As String

    XX = "ABCD" '选择题选项
    s = Len(XX) '取选项字符串长度


    For i = 1 To ActiveDocument.Paragraphs.count

        ActiveDocument.Range(ActiveDocument.Paragraphs(i).Range.Start, ActiveDocument.Paragraphs(i).Range.End).Select    '选定本段

        If InStrRev(ActiveDocument.Paragraphs(i).Range.text, "【　　】") Then
      
            Selection.Find.ClearFormatting
            Selection.Find.MatchWildcards = True

            For m = 1 To Int(s / 2)
                Randomize (Timer)

                x = Int((4 * Rnd()) + 1) '生成1-4的随机数

                ActiveDocument.Range(ActiveDocument.Paragraphs(i).Range.Start, ActiveDocument.Paragraphs(i).Range.End).Select    '选定本段

                'MsgBox Mid(XX, x, 1)

                Selection.Find.Execute findtext:=Mid(XX, x, 1) & "[．.、]"  'mid函数：返回文本字符串中从指定位置开始的特定数目的字符，每次取一个数字

                Selection.MoveRight Unit:=wdCharacter, count:=1     '光标右移

                Ts = Selection.Start    '取光标所在为位置为选定起点TempStart

                Selection.Find.Execute findtext:="[^9^11^13]"

                Selection.MoveLeft Unit:=wdCharacter, count:=1     '光标移动

                Te = Selection.Start    '取光标所在为位置为选定终点TempEnd

                Selection.Start = Ts
                Selection.End = Te
                
                If Ts = Te Then GoTo k

                Selection.Cut

                Y = Int((4 * Rnd()) + 1) '生成1-4的随机数

                If Y = x Then Y = Int((4 * Rnd()) + 1)            '如果y=x，再生成一个1-4的随机数

                ActiveDocument.Range(ActiveDocument.Paragraphs(i).Range.Start, ActiveDocument.Paragraphs(i).Range.End).Select    '选定本段

                Selection.Find.Execute findtext:=Mid(XX, Y, 1) & "[．.、]"  'mid函数：返回ABCD中从指定位置开始的特定数目的字符，每次取1个数字

    '                Selection.Find.Execute findtext:="[A-D][．.、]"   '在已经选定部分查找并选定
    '                Selection.MoveLeft unit:=wdCharacter, Count:=1     '光标左移
                Selection.MoveRight Unit:=wdCharacter, count:=1     '光标右移

                Selection.Paste          '粘贴第1次剪切内容

                Ts = Selection.Start    '取光标所在为位置为选定起点TempStart

                Selection.Find.Execute findtext:="[^9^11^13]"

                Selection.MoveLeft Unit:=wdCharacter, count:=1     '光标移动

                Te = Selection.Start    '取光标所在为位置为选定终点TempEnd

                Selection.Start = Ts
                Selection.End = Te

                If Ts = Te Then GoTo k

                Selection.Cut        '第2次剪切

                ActiveDocument.Range(ActiveDocument.Paragraphs(i).Range.Start, ActiveDocument.Paragraphs(i).Range.End).Select    '选定本段

                Selection.Find.Execute findtext:=Mid(XX, x, 1) & "[．.、]"  '本段第一次剪切的选项选定

                'Selection.Find.Execute findtext:="[A-D][．.、]"   '在已经选定部分查找并选定

                Selection.MoveRight Unit:=wdCharacter, count:=1     '光标右移到第一次剪切的选项后
                Selection.Paste             '粘贴第2次剪切内容
k:
            Next m
            n = n + 1
        Else
        End If
    Next i

    '    If n Then
    '        MsgBox "共有" & n & "道选择题选项调整。"
    '    Else
    '        MsgBox "没有选择题选项调整。"
    '
    '    End If
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = " "
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
    '        .Format = True
    '        .MatchCase = False
    '        .MatchWholeWord = False
    '        .MatchByte = False
    '        .MatchWildcards = False
    '        .MatchSoundsLike = False
    '        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    '    Application.ScreenUpdating = True '屏幕更新


End Sub

Sub 格式化选择题()



      '以下数字变量的单位均为“磅”
    
    ActiveDocument.ActiveWindow.View.ShowAll = False    '隐藏所有编辑标记

    Application.ScreenUpdating = False '关闭屏幕更新



    Dim msgResult As VbMsgBoxResult

    msgResult = MsgBox("将格式化选择题，使4个选项对齐。" & Chr(13) & Chr(13) _
                & " 点“是”,则继续， 点“否”,则退出。" & Chr(13) & Chr(13) _
                & " 由于试卷的版面比较复杂，个别内容需要手动调整。", vbYesNo)
    Select Case msgResult

        Case vbYes
        
        查找替换合并选择项重新分割
        
        Selection.WholeStory
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = ""
                .Replacement.text = ""
                .Forward = True
                .Wrap = wdFindContinue
    
            End With
        Selection.Find.Execute Replace:=wdReplaceAll
 
        Selection.HomeKey Unit:=wdStory    ' 移动光标至文档开始
        
        Selection.Find.Execute findtext:="【　　】"  '查找并选定

        sj = Selection.ParagraphFormat.LeftIndent
        
    '         MsgBox sj

        Case vbNo '如果用户选择“否”
        
        Exit Sub '退出
    End Select

    '    Application.Run MacroName:="查找替换合并选择项重新分割"

    Dim Width As Single, Width2 As Single, Left As Single, Right As Single
    
    Width2 = ActiveDocument.PageSetup.TextColumns.Width '分栏文字宽度，只适合等宽分栏
    
    Tab_L = Int((Width2 - sj) / 4) '相邻制表位的宽度，取整数
    '
    '    Selection.Find.ClearFormatting
    '    Selection.Find.Replacement.ClearFormatting
    '    With Selection.Find
    '        .text = ""
    '        .Replacement.text = ""
    '        .Forward = True
    '        .Wrap = wdFindContinue
    '
    '    End With
    '    Selection.Find.Execute Replace:=wdReplaceAll

    '为整个文件页面设置选择题项的5个制表位

    Selection.WholeStory
    Selection.ParagraphFormat.TabStops.ClearAll
    '    ActiveDocument.DefaultTabStop = sj
    Selection.ParagraphFormat.TabStops.Add Position:=sj, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    Selection.ParagraphFormat.TabStops.Add Position:=sj + Tab_L, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    Selection.ParagraphFormat.TabStops.Add Position:=(sj + 2 * Tab_L), Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    Selection.ParagraphFormat.TabStops.Add Position:=(sj + 3 * Tab_L), Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    Selection.ParagraphFormat.TabStops.Add Position:=Width2, Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces

    'Exit Sub

    '    Dim a, B, C, D, n, m, Mp, t As Single


    For i = 1 To ActiveDocument.Paragraphs.count
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = ""
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue

        End With
        Selection.Find.Execute Replace:=wdReplaceAll

        ActiveDocument.Range(ActiveDocument.Paragraphs(i).Range.Start, ActiveDocument.Paragraphs(i).Range.End).Select    '选定本段
        
        ' Selection.MoveDown Unit:=wdLine, Count:=1
        
        Selection.Find.Execute findtext:="【　　】"  '查找并选定

        If InStrRev(ActiveDocument.Paragraphs(i).Range.text, "【　　】") Then
        
    '        sj = Selection.ParagraphFormat.LeftIndent

            ' Selection.MoveDown Unit:=wdLine, Count:=1

            Selection.Find.Execute findtext:="A．"  '查找并选定

            Selection.EndKey Unit:=wdLine     ' 移动光标至当前行尾

            a1 = Selection.Information(wdHorizontalPositionRelativeToTextBoundary) '返回所选内容相对于周围最近的正文边界左边缘的水平位置，如果所选内容或区域未处于屏幕区域中，则该参数返回 - 1。
            a1 = a1 - sj
            'MsgBox a1

            ActiveDocument.Range(ActiveDocument.Paragraphs(i).Range.Start, ActiveDocument.Paragraphs(i).Range.End).Select    '选定本段

            Selection.Find.Execute findtext:="B．"  '查找并选定
            'Selection.MoveDown Unit:=wdLine, Count:=1
            Selection.EndKey Unit:=wdLine     ' 移动光标至当前行尾

            b1 = Selection.Information(wdHorizontalPositionRelativeToTextBoundary) '返回所选内容相对于周围最近的正文边界左边缘的水平位置，如果所选内容或区域未处于屏幕区域中，则该参数返回 - 1。

            b1 = b1 - sj
            ' MsgBox b1

            ActiveDocument.Range(ActiveDocument.Paragraphs(i).Range.Start, ActiveDocument.Paragraphs(i).Range.End).Select    '选定本段


            Selection.Find.Execute findtext:="C．"  '查找并选定

            ' Selection.MoveDown Unit:=wdLine, Count:=1
            Selection.EndKey Unit:=wdLine     ' 移动光标至当前行尾

            c1 = Selection.Information(wdHorizontalPositionRelativeToTextBoundary) '返回所选内容相对于周围最近的正文边界左边缘的水平位置，如果所选内容或区域未处于屏幕区域中，则该参数返回 - 1。
            c1 = c1 - sj

            'MsgBox c1

            ActiveDocument.Range(ActiveDocument.Paragraphs(i).Range.Start, ActiveDocument.Paragraphs(i).Range.End).Select    '选定本段


            Selection.Find.Execute findtext:="D．"  '查找并选定
            'Selection.MoveDown Unit:=wdLine, Count:=1
            Selection.EndKey Unit:=wdLine     ' 移动光标至当前行尾

            d1 = Selection.Information(wdHorizontalPositionRelativeToTextBoundary) '返回所选内容相对于周围最近的正文边界左边缘的水平位置，如果所选内容或区域未处于屏幕区域中，则该参数返回 - 1。
            d1 = d1 - sj
            ' MsgBox d1

            ActiveDocument.Range(ActiveDocument.Paragraphs(i).Range.Start, ActiveDocument.Paragraphs(i).Range.End).Select    '选定本段



            m = a1
            If b1 > m Then m = b1   '求四个选项的最大值
            If c1 > m Then m = c1
            If d1 > m Then m = d1
            'MsgBox M
            'Exit Sub

            m = m - sj

            If m < Tab_L Then
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                With Selection.Find
                    .text = "^11([BCD]．)"
                    .Replacement.text = "^9\1"
                    .Forward = False
                    .Wrap = wdFindStop
                    .MatchWildcards = True
                End With
                Selection.Find.Execute Replace:=wdReplaceAll

            Else

                If m < Tab_L * 2 Then
                    Selection.Find.ClearFormatting
                    Selection.Find.Replacement.ClearFormatting

                    With Selection.Find
                        .text = "^11B．"
                        If a1 < Tab_L Then
                            .Replacement.text = "^9^9B．"
                        Else
                            .Replacement.text = "^9B．"
                        End If

                        .Forward = False
                        .Wrap = wdFindStop
                        .MatchWildcards = True
                    End With
                    Selection.Find.Execute Replace:=wdReplaceAll

                    With Selection.Find
                        .text = "^11D．"
                        If c1 < Tab_L Then
                            .Replacement.text = "^9^9D．"
                        Else
                            .Replacement.text = "^9D．"
                        End If

                        .Forward = False
                        .Wrap = wdFindStop
                        .MatchWildcards = True
                    End With
                    Selection.Find.Execute Replace:=wdReplaceAll
                Else

                End If

            End If


            Dim TAB_n As Integer
            ActiveDocument.Range(ActiveDocument.Paragraphs(i).Range.Start, ActiveDocument.Paragraphs(i).Range.End).Select

            '清空查找替换框
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = ""
                .Replacement.text = ""
                .Forward = True
                .Wrap = wdFindContinue

            End With
            Selection.Find.Execute Replace:=wdReplaceAll

            Selection.Find.Execute findtext:="【　　】"

            Line_KD = Selection.Information(wdHorizontalPositionRelativeToTextBoundary) '返回所选内容相对于周围最近的正文边界左边缘的水平位置，如果所选内容或区域未处于屏幕区域中，则该参数返回 - 1。

            zihao = Selection.Font.Size      '字号

            Line_SY = Width2 - Line_KD          '剩余宽度

            TAB_n = -Int(-Line_SY / Tab_L)  '向上取整，计算应当添加的制表符个数

            ' If Line_SY / Tab_L = 1 Or Line_SY / Tab_L = 2 Or Line_SY / Tab_L = 3 Then TAB_n = TAB_n + 1

            ActiveDocument.Range(ActiveDocument.Paragraphs(i).Range.Start, ActiveDocument.Paragraphs(i).Range.End).Select
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "【　　】"
                Select Case TAB_n
                    Case 0
                        .Replacement.text = "^11^9^9^9^9【　　】"

                    Case 1

                        If Line_SY > 4 * zihao Then

                            .Replacement.text = "^9【　　】"

                        Else

                            .Replacement.text = "【　　】"

                        End If

                    Case 2
                        .Replacement.text = "^9^9【　　】"

                    Case 3
                        .Replacement.text = "^9^9^9【　　】"

                    Case 4
                        .Replacement.text = "^9^9^9^9【　　】"
                    Case 5
                        'If SJ <> 0 Then
                        .Replacement.text = "^11^9^9^9^9【　　】"
                        'Else
                        '   .Replacement.Text = "^11^9^9^9^9^9【　　】"
                        'End If
                End Select
                .Forward = False
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchByte = False
                .MatchAllWordForms = False
                .MatchSoundsLike = False
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll

            If a1 * b1 * c1 * d1 Then

                n = n + 1

                ' MsgBox "第" & i & "段，选择项最长为" & M & "厘米"

                ' Else

                ' MsgBox "第" & i & "段不是选择题。"

            End If
        Else

        End If

E:

    Next i
    
    Selection.MoveUp Unit:=wdParagraph


    If n Then
        'MsgBox "有选择题。"
    Else
        MsgBox "本文件没有选择题，或者选择项没有整理到一个段落。"

    End If
    ActiveDocument.ActiveWindow.View.ShowAll = True   '不隐藏所有编辑标记

Application.ScreenUpdating = True '屏幕更新
 
End Sub

Sub 查找替换合并选择项重新分割()

    'Sub 查找替换合并选择项()
    ' 此宏有另外一个宏“格式化选择题”引用，不得修改宏名。
    '
  Selection.HomeKey Unit:=wdStory

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(^13){1,}"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll



    '查找BCD选项前的空格、制表符、分行符和回车符，替换为分行符，保留BCD，并把BCD后的标点统一



    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "[ 　^9^11^13]{1,}([A-D])[.．、。]{1,}"
        .Replacement.text = "^11\1．"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
     '清除括号格，统一括号格式
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "[\(（【][\)）】][^11^13]A[.．、]"
        .Replacement.text = "【　　】^11A．"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
 '清除括号格，统一括号格式
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "[\(（【][ 　]{1,}[\)）】][^11^13]A[.．、]"
        .Replacement.text = "【　　】^11A．"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    
    
    '解决题干与选择支不连续情况（主要是有图）
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "[\(（][ 　]{1,}[\)）][^11^13]"
        .Replacement.text = "【　　】^11"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     '清除括号前的空格制表符
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "[ 　^9]{1,}(【　　】^11A．)"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
        
         '清除重复执行的括号前的制表符
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "[^9]{1,}(【　　】)"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    
         '清除重复执行的括号前的手动换行及制表符
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "^11[^9]{1,}(【　　】)"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    

    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "^13([图])"
        .Replacement.text = "^l\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
 
    
End Sub

Sub 查找替换合并选择项()

    '
    ' 此宏有另外一个宏“格式化选择题”引用，不得修改宏名。
    '
    '
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "[ ^9^11^13]{1,}([B-D])[.．、]"
        .Replacement.text = "^9\1．"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub 隐藏选择题答案()
    'Attribute 隐藏选择题答案.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.宏1"
    '
    '样式例子
    '(2017·武汉)下列事例中利用声传递能量的是(C)        【　　】
    'A.通过声学仪器接收到的次声波判断地震的方位
    'B.利用超声导盲仪探测前进道路上的障碍物
    'C.利用超声波排除人体内的结石
    'D.利用超声波给金属工件探伤

    '
    '

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Hidden = True
        .Color = wdColorRed
    End With
    With Selection.Find
        .text = "\([A-D]{1,2}\)^9"
        .Replacement.text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Hidden = True
        .Color = wdColorRed
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Hidden = False
        .Color = wdColorAutomatic
    End With
    With Selection.Find
        .text = "^9"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'Application.Run MacroName:="填空题下划线空格替换为全角下划线字符"
    
End Sub



Sub 将答案统一挪到文档末尾()

    '题目如下:
    '1.头孢菌素类药物是.A
    'a.抗菌谱较广的抗生素
    'B.治疗非典型肺炎的药物
    'C.抗真菌药
    'D.抗结核病药
    'E.抗病毒药
    '2.联合应用抗菌药物的目的不包括B
    'a.发挥药物的协同作物以提高疗效
    'B.使医生和病人都有一种安全感
    'C.延迟或减少耐药菌株的产生
    'D.对混合感染扩大抗菌范围
    'E.减少个别药剂量，以减少毒副反应

ActiveDocument.Range.InsertAfter "整理的答案是："
With ActiveDocument.Content.Find
    .text = "([0-9]{1,2})*([A-Z])^13"
    .Forward = True
    .MatchWildcards = True
Do While .Execute
ActiveDocument.Range.InsertAfter Mid(.Parent.text, Len(.Parent.text) - 1, 1)
ActiveDocument.Range(.Parent.End - 2, .Parent.End - 1) = ""
.Parent.Collapse Direction:=wdCollapseEnd
Loop
End With
End Sub

Sub 将答案统一挪到文档末尾的逆向过程()


    '将答案统一挪到文档末尾的逆向过程
    On Error GoTo exitsub
    '假定最后一段为答案，如你所描述的那样1、A  2、B  3、C  4、C  5、B  6、E  7、C
    Dim s, i As Integer, p As Range
    Set p = ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.count).Range
    s = Split(p.text, "、")
    p.Delete
    With ActiveDocument.Content.Find
        .text = "([0-9]{1,2}).*^13"
        .Forward = True
        .MatchWildcards = True
        Do While .Execute
            i = i + 1
            If i > UBound(s) Then MsgBox "题目与答案不一致，将退出"
            ActiveDocument.Range(.Parent.End - 1, .Parent.End - 1).InsertAfter Left(s(i), 1)
            .Parent.Collapse Direction:=wdCollapseEnd
        Loop
    End With
exitsub:
End Sub

Sub 选择题文档尾段答案逐一写入题干()
    'http://club.excelhome.net/thread-1032236-1-1.html
    Dim A, i As Integer
    ActiveDocument.Content.Find.Execute findtext:="^32", replacewith:="", Replace:=wdReplaceAll
    A = Split(ActiveDocument.Paragraphs.Last.Range, "、")
    'ActiveDocument.Paragraphs.Last.Range.Delete
    With ActiveDocument.Content.Find
        .text = "[\(（][\)）]"
        .Forward = True
        .MatchWildcards = True
        Do While .Execute
            i = i + 1
            If i > UBound(A) Then
                MsgBox "题目与答案不一致，将退出"
                Exit Sub
            End If
            ActiveDocument.Range(.Parent.End - 1, .Parent.End - 1).InsertAfter Left(A(i), 1)
            .Parent.Collapse Direction:=wdCollapseEnd
        Loop
    End With
End Sub

Sub 选择题文档尾段答案逐一写入题干2()
Dim reg As Object, j%
Set reg = CreateObject("vbscript.regexp")
With reg
       .Global = True
       .Pattern = "[A-Z]+(?=[\d+\s])"
       .MultiLine = True
       Set matches = .Execute(ActiveDocument.Paragraphs.Last.Range)
End With
Set reg = Nothing
With CreateObject("vbscript.regexp")
       .Global = True
       .MultiLine = True
j = 0
      Do While j < matches.count
           For Each para In ActiveDocument.Paragraphs
               .Pattern = "[（/(]\s*[）/)]"
               If .test(para.Range) = True Then
                    para.Range = .Replace(para.Range, "（" & matches.Item(j) & "）")
                    j = j + 1
               End If
           Next
       Loop
End With
End Sub


Sub 选择题末尾的答案提取到题干后的括号内()
    '
    '要求：将选枝下面【答案】后面的答案项，移到题干后面的括号里。然后删去原答案所在的行。
    '
    '
    '1. （2004年绵阳市）下列各项中，朗读的奏划分正确的一项是（ ）
    'A.故乡的歌/是一支/清远的笛，总/在有月亮的晚上/响起
    'B.开轩/面/场圃，把酒/话/桑麻
    'C.故/天/将降大任/于是人也
    'D.有的人/活着/他/已经/死了；有的人/死了/他/还活着
    '【答案】D
    '


 Dim st, en
    With ThisDocument.Range.Find
        Do While .Execute("[0-9]{1,}*【答案】*^13", , , 1)
            With .Parent
                .Collapse
                .MoveUntil Chr(13)
                .MoveUntil "（(", wdBackward
                .MoveEndUntil ")）"
                st = .Start: en = .End
                .Collapse 0
                .MoveUntil "【"
                .MoveUntil Chr(13)
                .MoveStartUntil "】", wdBackward
                ThisDocument.Range(st, en).FormattedText = .FormattedText
                .Expand 4
                .text = Empty
            End With
        Loop
    End With
End Sub

Sub 选择题单选题各题答案按题号提取到文档末尾()
     Dim i%
     ActiveDocument.Content.Find.Execute findtext:="^32", replacewith:="", Replace:=wdReplaceAll
     ActiveDocument.Range.InsertAfter Chr(13) & "整理的答案是："
     With ActiveDocument.Content.Find
         .text = "[\(（][A-Z]{1,}"
         .Forward = True
         .MatchWildcards = True
         Do While .Execute
             i = i + 1
             ActiveDocument.Range.InsertAfter i & "、" & VBA.Right(.Parent.text, Len(.Parent.text) - 1) & vbTab
             ActiveDocument.Range(.Parent.Start + 1, .Parent.End) = ""
             .Parent.Collapse Direction:=wdCollapseEnd
         Loop
     End With
End Sub



Sub 提取单选题答案按题号到文档末()
    'http://club.excelhome.net/thread-1032236-1-1.html
    Dim i%
    ActiveDocument.Content.Find.Execute findtext:="^32", replacewith:="", Replace:=wdReplaceAll
    ActiveDocument.Range.InsertAfter Chr(13) & "整理的答案是："
    With ActiveDocument.Content.Find
        .text = "[\(（][A-Z]"
        .Forward = True
        .MatchWildcards = True
        Do While .Execute
            i = i + 1
            ActiveDocument.Range.InsertAfter i & "、" & Right(.Parent.text, 1) & vbTab
            ActiveDocument.Range(.Parent.End - 1, .Parent.End) = ""
            .Parent.Collapse Direction:=wdCollapseEnd
        Loop
    End With
End Sub

Sub 提取答案到题目末尾()

    '求提取答案 ，将答案提取到题后，用括号存起来
    'http://club.excelhome.net/thread-1396523-1-1.html

    Dim myStart&, myDoc As Document, B As Boolean, Q As Range, R As Range, sr$
    Application.ScreenUpdating = False
    Set myDoc = ActiveDocument
    With myDoc.Content.Find
        Do While .Execute("^13[0-9]@[.、．]", , , -1, , , 0)
            With .Parent
                If Not B Then
                    Set Q = myDoc.Range(.Start + 1, myDoc.Content.End)
                    Set R = Q.Duplicate
                    With R.Find
                        .Font.ColorIndex = wdRed
                        Do While .Execute("*", , , -1)
                            If Not R.InRange(Q) Then Exit Do
                            sr = sr & R.text
                        Loop
                        With Q
                            If sr <> "" Then
                                .End = .End - 1: .InsertAfter "（" & sr & "）": sr = Empty
                            End If
                        End With
                    End With
                    B = True
                Else
                    Set Q = myDoc.Range(.Start + 1, myStart)
                    Set R = Q.Duplicate
                    With R.Find
                        .Font.ColorIndex = 6
                        Do While .Execute("*", , , -1)
                            If Not R.InRange(Q) Then Exit Do
                            sr = sr & R.text
                        Loop
                        With Q
                            If sr <> "" Then
                                .End = .End - 1: .InsertAfter "（" & sr & "）": sr = Empty
                            End If
                        End With
                    End With
                End If
                myStart = .Start + 1: .Collapse
            End With
        Loop
    End With
    Application.ScreenUpdating = True
    MsgBox "ok!"
End Sub
