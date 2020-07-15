Attribute VB_Name = "LayoutMacros"
Sub GetLayout()
Attribute GetLayout.VB_Description = "画像や枠線のレイアウトを CSS として取得する。"
Attribute GetLayout.VB_ProcData.VB_Invoke_Func = "Normal.LayoutMacros.GetLayout"
'
' GetLayout Macro
' 画像や枠線のレイアウトを CSS として取得する。
' Get layout of images and borders as CSS
'
    Dim mm, pw, wh, pl, pt As Single
    
    mm = 2.83464 '[pt/mm]
    
    ' ページのサイズを取得
    ' Get page size
    With ActiveDocument.ActiveWindow.Panes(1).Pages.Item(1)
        pw = .Width / mm '[mm]
        ph = .Height / mm '[mm]
        Msg = Msg + ".body {" + vbCrLf
        Msg = Msg + "    position: absolute;" + vbCrLf
        Msg = Msg + "    width: " + Format(pw, "0.00") + "mm;" + vbCrLf
        Msg = Msg + "    height: " + Format(ph, "0.00") + "mm;" + vbCrLf
        Msg = Msg + "    left: " + Format(.Left / mm, "0.00") + "mm;" + vbCrLf
        Msg = Msg + "    top: " + Format(.Top / mm, "0.00") + "mm;" + vbCrLf
        Msg = Msg + "    --content: <p>""body""</p>;" + vbCrLf
        Msg = Msg + "}" + vbCrLf + vbCrLf
    End With
    
    ' ページの余白を取得
    ' Get page margines
    With ActiveDocument.PageSetup
        pl = .LeftMargin / mm '[mm]
        pt = .TopMargin / mm '[mm]
    End With
    
    ' 画像/枠線のサイズを取得
    ' Get images/borders sizes
    For I = ActiveDocument.Shapes.Count To 1 Step -1
        With ActiveDocument.Shapes(I)
            Msg = Msg + "." + .Name + " {" + vbCrLf
            Msg = Msg + "    position: absolute;" + vbCrLf
            Msg = Msg + "    width: " + Format(.Width / mm / pw * 100, "0.00") + "%;" + vbCrLf
            Msg = Msg + "    height: " + Format(.Height / mm / ph * 100, "0.00") + "%;" + vbCrLf
            Msg = Msg + "    left: " + Format((.Left / mm + pl) / pw * 100, "0.00") + "%;" + vbCrLf
            Msg = Msg + "    top: " + Format((.Top / mm + pt) / ph * 100, "0.00") + "%;" + vbCrLf
            If .Type = msoTextBox Then
                Dim r As Byte, g As Byte, b As Byte
                r = .Line.ForeColor \ 256 ^ 1 Mod 256
                g = .Line.ForeColor \ 256 ^ 1 Mod 256
                b = .Line.ForeColor \ 256 ^ 2 Mod 256

                bc = "rgb(" & r & "," & g & "," & b & ")"
                Msg = Msg + "    border-color: " + bc + ";" + vbCrLf
                Msg = Msg + "    border-width: " + Format(.Line.Weight, "0.00") + "pt;" + vbCrLf
            End If
            Msg = Msg + "    --content: <p>""" + .Name + """</p>;" + vbCrLf
            Msg = Msg + "}" + vbCrLf + vbCrLf
        End With
    Next
    
    ' 新しいテキストボックスを作成して結果をセット
    ' Create a new text box and set the result
    With ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        10, 10, 300, 600)
        .Name = "layout"
        .TextFrame.TextRange.Font.Size = 10
        .TextFrame.TextRange.Font.Name = "Meiryo UI"
        .TextFrame.TextRange.Font.Color = RGB(100, 100, 100)
        .TextFrame.TextRange.Text = Msg
    End With
End Sub
Sub ClearLayout()
Attribute ClearLayout.VB_Description = "作成したレイアウトを削除。"
Attribute ClearLayout.VB_ProcData.VB_Invoke_Func = "Normal.LayoutMacros.ClearLayout"
'
' ClearLayout Macro
' 作成したレイアウトを削除
'
    For I = ActiveDocument.Shapes.Count To 1 Step -1
        With ActiveDocument.Shapes(I)
            If .Name = "layout" Then
                .Delete
            End If
        End With
    Next
End Sub
