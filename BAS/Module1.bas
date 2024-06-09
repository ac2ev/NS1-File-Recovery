Attribute VB_Name = "Module1"
Option Explicit

Public Sub ListviewPrint(ByVal Title As String, ByVal Listview As Listview, Optional ByVal LeftMargin As Long = 400, Optional ByVal TopMargin As Long = 400)

  Dim i As Integer
  Dim intPageNumber As Integer

  Dim itmX As ListItem

  Dim lngYMargin As Long
  Dim iCHWidth As Long

  Dim strDate As String

  strDate = "(Printed on " & Format$(Now, "Long Date") & ")"

' Print the Title
  With Printer
    With .Font
      .Bold = True
      .Italic = False
      .Name = Listview.Font.Name
      .Size = 14
    End With
    .CurrentY = TopMargin + Printer.TextHeight("Xg")
    .CurrentX = (Printer.width - Printer.TextWidth(Title)) / 2
    Printer.Print Title
    With .Font
      .Bold = False
      .Italic = True
      .Size = 12
    End With
    Printer.Print
    .CurrentX = LeftMargin
    .CurrentX = (Printer.width - Printer.TextWidth(strDate)) / 2
    Printer.Print strDate
    Printer.Print
  End With

' Print the column Headings
  ListviewPrintColumnHeaders Listview, LeftMargin

  For Each itmX In Listview.ListItems
    With Printer
      With .Font
        .Bold = False
        .Italic = False
        .Size = 10
      End With
      If .CurrentY > _
        .Height - 2 * TopMargin - 3 * .TextHeight("Xg") Then
        ListviewPrintPageNumber intPageNumber, TopMargin
        .NewPage
        .CurrentY = TopMargin + Printer.TextHeight("Xg")
        .CurrentX = (Printer.width - Printer.TextWidth(Title & " [Cont.]"))
/ 2
        Printer.Print Title & " [Cont.]"
        Printer.Print
        ListviewPrintColumnHeaders Listview, LeftMargin
      End If
      With .Font
        .Bold = False
        .Italic = Listview.Font.Italic
        .Size = Listview.Font.Size
      End With
      lngYMargin = .CurrentY
      iCHWidth = 0
      If Listview.ColumnHeaders(1).width > 0 Then
        With Listview.ColumnHeaders(1)
          If .Alignment = lvwColumnleft Then
            Printer.CurrentX = LeftMargin + iCHWidth
          ElseIf .Alignment = lvwColumncenter Then
            Printer.CurrentX = LeftMargin + iCHWidth + (.width _
              - Printer.TextWidth(.Text)) / 2
          ElseIf .Alignment = lvwColumnright Then
            Printer.CurrentX = LeftMargin + iCHWidth + .width _
              - Printer.TextWidth(.Text)
          End If
          Printer.Print itmX.Text
          Printer.CurrentY = lngYMargin
          iCHWidth = iCHWidth + .width
        End With
      End If
      For i = 2 To Listview.ColumnHeaders.Count
        If Listview.ColumnHeaders(i).width > 0 Then
          With Listview.ColumnHeaders(i)
            If .Alignment = lvwColumnleft Then
              Printer.CurrentX = LeftMargin + iCHWidth
            ElseIf .Alignment = lvwColumncenter Then
              Printer.CurrentX = LeftMargin + iCHWidth + (.width _
                - Printer.TextWidth(.Text)) / 2
            ElseIf .Alignment = lvwColumnright Then
              Printer.CurrentX = LeftMargin + iCHWidth + .width _
                - Printer.TextWidth(.Text)
            End If
            Printer.Print itmX.SubItems(i - 1)
            Printer.CurrentY = lngYMargin
            iCHWidth = iCHWidth + .width
          End With
        End If
      Next i
      .CurrentY = .CurrentY + _
        .TextHeight("Xg")
    End With
  Next itmX

  With Printer
    ListviewPrintPageNumber intPageNumber, TopMargin
    .EndDoc
  End With

End Sub

Private Sub ListviewPrintColumnHeaders(ByVal Listview As Listview, Optional ByVal LeftMargin As Long = 100)

  Dim chdColumn As ColumnHeader

  Dim lngYMargin As Long
  Dim iCHWidth As Long

  With Printer.Font
    .Bold = True
    .Italic = Listview.Font.Italic
    .Size = Listview.Font.Size
  End With
  lngYMargin = Printer.CurrentY
  For Each chdColumn In Listview.ColumnHeaders
    With chdColumn
      If .width > 0 Then
        Printer.Line (LeftMargin + iCHWidth, Printer.CurrentY + Printer.TextHeight("Mg")) _
          -(LeftMargin + iCHWidth + .width, Printer.CurrentY + Printer.TextHeight("Mg"))

        If .Alignment = lvwColumnleft Then
          Printer.CurrentX = LeftMargin + iCHWidth
        ElseIf .Alignment = lvwColumncenter Then
          Printer.CurrentX = LeftMargin + iCHWidth + (.width _
            - Printer.TextWidth(.Text)) / 2
        ElseIf .Alignment = lvwColumnright Then
          Printer.CurrentX = LeftMargin + iCHWidth + .width _
            - Printer.TextWidth(.Text)
        End If

        Printer.Print .Text
        Printer.Line (LeftMargin + iCHWidth, Printer.CurrentY) _
          -(LeftMargin + iCHWidth + .width, Printer.CurrentY)

        Printer.CurrentY = lngYMargin
        iCHWidth = iCHWidth + .width
      End If
    End With
  Next chdColumn
  Printer.CurrentY = lngYMargin + 2 * Screen.TwipsPerPixelY + Printer.TextHeight("Xg")
  Printer.CurrentY = Printer.CurrentY + _
    Printer.TextHeight("Xg")

End Sub

Private Sub ListviewPrintPageNumber(ByRef PageNumber As Integer, ByValBottomMargin As Integer)

  With Printer
    With .Font
      .Bold = False
      .Italic = False
      .Size = 10
    End With
    PageNumber = PageNumber + 1
    .CurrentY = Printer.Height - _
      BottomMargin - 2 * Printer.TextHeight("Xg")
    .CurrentX = (Printer.width - Printer.TextWidth("Page " & PageNumber)) / 2
    Printer.Print "Page " & PageNumber
  End With

End Sub

