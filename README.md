# PowerPoint Font Changer Macro

## Purpose
This macro was created to batch-change the font of all text in a PowerPoint presentation to "MS Mincho."

## Tested Environment
Verified to work with Microsoft 365.

## How to Run the Macro

1. **Enable the Developer Tab**:
   - Open PowerPoint and go to `File` > `Options` > `Customize Ribbon`.
   - Check the box for `Developer` and click `OK`.

2. **Add the Macro**:
   - Click the `Developer` tab, then select `Macros` > `Visual Basic`.
   - Go to `Insert` > `Module` and paste the macro code provided below.
   - Save the file by selecting `File` > `Save`, choosing the `PowerPoint Macro-Enabled Presentation (*.pptm)` format.

3. **Trust Center Settings**:
   - Go to `File` > `Options` > `Trust Center` > `Trust Center Settings`.
   - Under `Macro Settings`, select `Enable all macros` (note: this poses a security risk, so revert this setting after use).
   - Under `Trusted Locations`, add the folder where the macro-enabled file is saved.

4. **Run the Macro**:
   - Click the `Developer` tab and select `Macros`.
   - Choose `ChangeFontToMSMincho` and click `Run`.

Following these steps will change the font of all text frames, tables, and grouped shapes in the PowerPoint presentation to "MS Mincho."

## Benefits
The font-changing process is encapsulated in a function, making it easy to modify for other fonts, which enhances maintainability.

## Macro Code

```vba
Option Explicit
Sub ChangeFontToMSMincho()
    ' アクティブなプレゼンテーション内のすべてのスライド、シェイプ、表、グループ化されたシェイプのフォントを
    ' 指定したフォントに変更するマクロ

    Dim sld As Slide
    Dim shp As Shape
    Dim tbl As Table
    Dim row As Integer
    Dim col As Integer
    Dim grpShp As Shape

    ' 変更したいフォント名
    Dim targetFont As String
    targetFont = "ＭＳ 明朝"

    On Error GoTo ErrHandler

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            ' 通常のテキストフレーム、図形内のテキストフレーム、表のセル内のテキスト
            If shp.HasTextFrame Then
                With shp.TextFrame.TextRange.Font
                    .NameFarEast = targetFont
                End With
            ElseIf shp.HasTable Then
                Set tbl = shp.Table
                For row = 1 To tbl.Rows.Count
                    For col = 1 To tbl.Columns.Count
                        With tbl.Cell(row, col).Shape.TextFrame.TextRange.Font
                            .NameFarEast = targetFont
                        End With
                    Next col
                Next row
            ElseIf shp.Type = msoGroup Then ' グループ化されたシェイプ
                For Each grpShp In shp.GroupItems
                    If grpShp.HasTextFrame Then
                        With grpShp.TextFrame.TextRange.Font
                            .NameFarEast = targetFont
                        End With
                    End If
                Next grpShp
            End If
        Next shp
    Next sld

Exit Sub

ErrHandler:
    MsgBox "フォントの変更中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー説明: " & Err.Description
End Sub
