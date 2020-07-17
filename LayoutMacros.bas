Attribute VB_Name = "LayoutMacros"
' LayoutMacros.bas v0.0.2
' Copyright (c) 2020 Nonki Takahashi.  The MIT License.
'
Sub GetLayout()
Attribute GetLayout.VB_Description = "Get layout of pictures and outlines as CSS."
Attribute GetLayout.VB_ProcData.VB_Invoke_Func = "Normal.LayoutMacros.GetLayout"
'
' GetLayout Macro
' Get layout of pictures and outlines as CSS
'
    Dim mm, pw, wh, pl, pt As Single
    
    mm = 2.83464 '[pt/mm]
    
    ' Get body style
    Msg = Msg + "body {" + vbCrLf
    Msg = Msg + "    background-color: lightgray;" + vbCrLf
    Msg = Msg + "    text-align: center;" + vbCrLf
    Msg = Msg + "    font-family: 'Meiryo UI';" + vbCrLf
    Msg = Msg + "    font-size: 10pt;" + vbCrLf
    Msg = Msg + "}" + vbCrLf + vbCrLf
    
    ' Get page margines
    With ActiveDocument.PageSetup
        pt = .TopMargin / mm '[mm]
        pr = .RightMargin / mm '[mm]
        pb = .BottomMargin / mm '[mm]
        pl = .LeftMargin / mm '[mm]
    End With
    
    ' Get page size
    With ActiveDocument.ActiveWindow.Panes(1).Pages.Item(1)
        pw = .Width / mm '[mm]
        ph = .Height / mm '[mm]
        Msg = Msg + ".page {" + vbCrLf
        Msg = Msg + "    background-color: white;" + vbCrLf
        Msg = Msg + "    position: relative;" + vbCrLf
        Msg = Msg + "    text-align: left;" + vbCrLf
        Msg = Msg + "    width: " + Format(pw, "0.000") + "mm;" + vbCrLf
        Msg = Msg + "    height: " + Format(ph, "0.000") + "mm;" + vbCrLf
        Msg = Msg + "    margin: 0 auto;" + vbCrLf
        Msg = Msg + "    padding-top: " + Format(pt, "0.000") + "mm;" + vbCrLf
        Msg = Msg + "    padding-right: " + Format(pr, "0.000") + "mm;" + vbCrLf
        Msg = Msg + "    padding-bottom: " + Format(pb, "0.000") + "mm;" + vbCrLf
        Msg = Msg + "    padding-left: " + Format(pl, "0.000") + "mm;" + vbCrLf
        ' display for CSS Preview
        Msg = Msg + "    --content: <p>""page""</p>;" + vbCrLf
        Msg = Msg + "}" + vbCrLf + vbCrLf
    End With
    
    
    ' Get pictures / outlines sizes
    For I = ActiveDocument.Shapes.Count To 1 Step -1
        With ActiveDocument.Shapes(I)
            Dim r As Byte, g As Byte, b As Byte
            Dim mt, mr, mb, ml As Single
            
            If .Type = msoTextBox Then
                r = .Line.ForeColor \ 256 ^ 1 Mod 256
                g = .Line.ForeColor \ 256 ^ 1 Mod 256
                b = .Line.ForeColor \ 256 ^ 2 Mod 256

                bc = "rgb(" & r & "," & g & "," & b & ")"
                mt = .TextFrame.MarginTop / mm '[mm]
                mr = .TextFrame.MarginRight / mm '[mm]
                mb = .TextFrame.MarginBottom / mm '[mm]
                ml = .TextFrame.MarginLeft / mm '[mm]
            Else
                mt = 0
                mr = 0
                mb = 0
                ml = 0
            End If
            Msg = Msg + "." + .Name + " {" + vbCrLf
            Msg = Msg + "    position: absolute;" + vbCrLf
            Msg = Msg + "    width: " + Format((.Width / mm - mr - ml) / pw * 100, "0.000") + "%;" + vbCrLf
            Msg = Msg + "    height: " + Format((.Height / mm - mt - mb) / ph * 100, "0.000") + "%;" + vbCrLf
            Msg = Msg + "    left: " + Format((.Left / mm + pl) / pw * 100, "0.000") + "%;" + vbCrLf
            Msg = Msg + "    top: " + Format((.Top / mm + pt) / ph * 100, "0.000") + "%;" + vbCrLf
            If .Type = msoTextBox Then
                bc = "rgb(" & r & "," & g & "," & b & ")"
                Msg = Msg + "    border-style: solid;" + vbCrLf
                Msg = Msg + "    border-color: " + bc + ";" + vbCrLf
                Msg = Msg + "    border-width: " + Format(.Line.Weight, "0.000") + "pt;" + vbCrLf
                Msg = Msg + "    padding-top: " + Format(mt, "0.000") + "mm;" + vbCrLf
                Msg = Msg + "    padding-right: " + Format(mr, "0.000") + "mm;" + vbCrLf
                Msg = Msg + "    padding-bottom: " + Format(mb, "0.000") + "mm;" + vbCrLf
                Msg = Msg + "    padding-left: " + Format(ml, "0.000") + "mm;" + vbCrLf
            End If
            ' display for CSS Preview
            Msg = Msg + "    --content: <p>""" + .Name + """</p>;" + vbCrLf
            Msg = Msg + "}" + vbCrLf + vbCrLf
        End With
    Next
    
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
Attribute ClearLayout.VB_Description = "Remove created layout."
Attribute ClearLayout.VB_ProcData.VB_Invoke_Func = "Normal.LayoutMacros.ClearLayout"
'
' ClearLayout Macro
' Remove created layout
'
    For I = ActiveDocument.Shapes.Count To 1 Step -1
        With ActiveDocument.Shapes(I)
            If .Name = "layout" Then
                .Delete
            End If
        End With
    Next
End Sub
