# Smart Data Label Arranger for Excel Line Charts

This VBA macro intelligently arranges data labels in an Excel line chart, ensuring labels are:
- Clearly readable
- Aligned to their X-axis positions
- Stacked vertically in multiple columns to avoid overlap
- Dynamically adjustable based on how many labels exist

## 📌 Features
- Automatically filters out empty labels
- Distributes labels into 8 vertical stacks (or fewer if needed)
- Aligns each label with its corresponding X-axis point
- Avoids clutter by stacking within each column
- Adjustable spacing and font size for readability

## 🧠 Use Case
When your Excel line chart has many overlapping data labels from cell values, this macro helps spread them out vertically in a structured, readable way while keeping them aligned to the data point.

## 💾 How to Use
1. Open your Excel workbook with a line chart.
2. Press `Alt + F11` to open the VBA Editor.
3. Insert a new module (`Insert > Module`) and paste the code below.
4. Close the editor and run the macro using `Alt + F8`.

## 🛠️ VBA Code
```vba
Sub ArrangeLabelsIn8StacksAlignedX_Improved()
    Dim ws As Worksheet
    Dim chtObj As ChartObject
    Dim srs As Series
    Dim i As Integer, groupIndex As Integer, posInGroup As Integer
    Dim offsetStep As Double
    Dim startTop As Double
    Dim labelsWithText As Collection
    Dim labelCount As Integer
    Dim groups As Integer
    Dim labelsPerGroup As Integer

    Set ws = ActiveSheet

    If ws.ChartObjects.Count = 0 Then
        MsgBox "No charts found on this sheet.", vbExclamation
        Exit Sub
    End If

    Set chtObj = ws.ChartObjects(1)
    Set srs = chtObj.Chart.SeriesCollection(1)

    If Not srs.HasDataLabels Then srs.ApplyDataLabels

    ' Collect labels with text only
    Set labelsWithText = New Collection
    For i = 1 To srs.Points.Count
        With srs.Points(i)
            If .HasDataLabel Then
                With .DataLabel
                    If Len(Trim(.Text)) > 0 Then
                        labelsWithText.Add i
                    Else
                        .Delete
                    End If
                End With
            End If
        End With
    Next i

    labelCount = labelsWithText.Count
    If labelCount = 0 Then
        MsgBox "No data labels with values found.", vbInformation
        Exit Sub
    End If

    groups = 8 ' Number of vertical stacks
    If labelCount < groups Then groups = labelCount

    labelsPerGroup = Int((labelCount + groups - 1) / groups) ' Ceiling division

    offsetStep = 20    ' Vertical spacing
    startTop = chtObj.Chart.PlotArea.InsideTop + 10  ' Starting top position

    ' Arrange labels
    For i = 1 To labelCount
        groupIndex = Int((i - 1) / labelsPerGroup)
        posInGroup = (i - 1) Mod labelsPerGroup

        With srs.Points(labelsWithText(i)).DataLabel
            .Position = xlLabelPositionAbove
            .Left = .Left
            .Top = startTop + posInGroup * offsetStep
            .Font.Size = 9
            .Orientation = 0
            .AutoScaleFont = True
        End With
    Next i

    MsgBox "Labels arranged for better visibility.", vbInformation
End Sub
