VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlResizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'0=Do nothing
'1=Proportional change
'2=Static change
'3=Limited Static change
Dim HPercent                As Double
Dim WPercent                As Double
Dim TPercent                As Double
Dim LPercent                As Double
Private FormMinWidth        As Long
Private FormMinHeight       As Long
Private FormOldWidth        As Long 'stores the old width of the owner form
Private FormOldHeight       As Long 'stores the old height of the owner form

Private Inited              As Boolean 'detect if resizer has been initiated
Private IgnoreNextIfNotMax  As Boolean 'ignore next resize if its not a maximise command

Public Sub InitResizer(frm As Form, MinWidth As Long, MinHeight As Long)
'this method will set the FormOldWidth and FormOldHeight
'and show that the resizer has been initiated
'Parameters:    frm: The owner form object
'Usage:         Simply add to Form_load event eg: frmResize.InitResizer(Me)

    FormOldWidth = frm.Width 'sets the current width of the form
    FormOldHeight = frm.Height 'sets the current height of the form
    FormMinWidth = MinWidth
    FormMinHeight = MinHeight
    
    Inited = True 'will show that the resizer has been initated
    
End Sub

Public Sub FormResized(frm As Form)
'this method will resize controls accordingly
'when the form is resized
'Parameters:    frm: The owner form object
'Usage:         simply add to form_resize event eg: Call frmResize.FormResized(Me)

    On Error GoTo Err_Handle 'error handler
    
    Dim cControl        As Control 'the variable holding a control from the owner form
    
    Dim WidthChange     As Long 'change in width of the owner form
    Dim HeightChange    As Long 'change in height of the owner form
    Dim WidthPercent    As Long 'percentage width change of the owner form
    Dim HeightPercent   As Long 'percentage height change of the owner form
    
    If Inited = False Then 'detect if InitResizer has been called
        'raise error if it hasnt
        Call Err.Raise(80000, , "Initiate resize module before resizing")
        Exit Sub 'stop further execution of this method to prevent errors
    End If
    
    If frm.WindowState = vbMinimized Then 'check to see if resize is a minimisation
        IgnoreNextIfNotMax = True 'set to ignore resize if its not to a maximise command
        Exit Sub 'stop further execution of this method to prevent errors
    End If
    
    If IgnoreNextIfNotMax = True Then 'detect if set to ignore next non-max command
        IgnoreNextIfNotMax = False 'set to perform resize next time
        If frm.WindowState <> vbMaximized Then 'will cause resize to display wrongly if not maximised
            Exit Sub 'stop further execution of this method to prevent errors
        End If
    End If

    'if new formsize is smaller then the specified minimumsize, then set to the minimum formsize
    If frm.Width < FormMinWidth Then
    frm.Width = FormMinWidth
    End If
    If frm.Height < FormMinHeight Then
    frm.Height = FormMinHeight
    End If
    
    WidthChange = frm.Width - FormOldWidth 'calculate width change
    HeightChange = frm.Height - FormOldHeight 'calculate height change
    
    WidthPercent = (((frm.Width - FormOldWidth) / FormOldWidth) * 100) 'calculate width percent change
    HeightPercent = (((frm.Height - FormOldHeight) / FormOldHeight) * 100) 'calculate height percent change
    
    For Each cControl In frm.Controls 'iterate through each control in the owner form
        
        'change left property
        Select Case Mid$(cControl.Tag, 1, 1) 'check left mode
            Case Is = 1 'Proportional change
                'change the left of control proportionally
                'cControl.Left = ((100 + WidthPercent) / 100) * cControl.Left
                LPercent = cControl.Left / FormOldWidth
                cControl.Left = frm.Width * LPercent
            Case Is = 2 'Static change
                cControl.Left = cControl.Left + WidthChange 'add change in width
            Case Is = 3 'Limited static change
                If cControl.Left + WidthChange < 0 Then 'detect if change will reduce below 0
                    cControl.Left = 0 'set left to 0
                Else
                    cControl.Left = cControl.Left + WidthChange 'add change as normal
                End If
        End Select
        
        'change top property
        Select Case Mid$(cControl.Tag, 2, 1) 'check top mode
            Case Is = 1 'Proportional change
                'change the top of control proportionally
                TPercent = cControl.Top / FormOldHeight
                cControl.Top = frm.Height * TPercent
                'cControl.Top = ((100 + HeightPercent) / 100) * cControl.Top
            Case Is = 2 'Static change
                cControl.Top = cControl.Top + HeightChange 'add change in height
            Case Is = 3 'Limited static change
                If cControl.Top + HeightChange < 0 Then 'detect if change will reduce below 0
                    cControl.Top = 0 'set top to 0
                Else
                    cControl.Top = cControl.Top + HeightChange 'add change as normal
                End If
        End Select
            
        'change width property
        Select Case Mid$(cControl.Tag, 3, 1) 'check width mode
            Case Is = 1 'Proportional change
                'change the left of control proportionally
                WPercent = cControl.Width / FormOldWidth
                cControl.Width = frm.Width * WPercent
            Case Is = 2 'Static change
                cControl.Width = cControl.Width + WidthChange 'add change in width
            Case Is = 3 'Limited static change
                If cControl.Width + WidthChange < 0 Then 'detect if change will reduce below 0
                    cControl.Width = 0 'set left to 0
                Else
                    cControl.Width = cControl.Width + WidthChange 'add change as normal
                End If
        End Select
        
        'change height property
        Select Case Mid$(cControl.Tag, 4, 1) 'check height mode
            Case Is = 1 'Proportional change
                'change the height of control proportionally
                HPercent = cControl.Height / FormOldHeight
                'cControl.Height = ((100 + HeightPercent) / 100) * cControl.Height
                cControl.Height = frm.Height * HPercent
            Case Is = 2 'Static change
                cControl.Height = cControl.Height + HeightChange 'add change in height
            Case Is = 3 'Limited static change
                If cControl.Height + HeightChange < 0 Then 'detect if change will reduce below 0
                    cControl.Height = 0 'set top to 0
                Else
                    cControl.Height = cControl.Height + HeightChange 'add change as normal
                End If
        End Select
        
    Next cControl 'goto next control
    
    FormOldWidth = frm.Width 'change old width to current width
    FormOldHeight = frm.Height 'change old height to current height

    Exit Sub

Err_Handle:
    If Err.Number = 438 Then '"object doesn't support this property or method" will occur on controls like timers
        Err.Clear 'clear error
        Resume Next 'ignore line with error and continue
    ElseIf Err.Number = 11 Then '"division by zero" will occur after restoring from minimize
        Err.Clear 'clear error
        Resume Next 'ignore line with error and continue
    ElseIf Err.Number = 80000 Then
        Exit Sub
    ElseIf Err.Number = 5 Then
        Exit Sub
    Else
        'show error box
        MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbExclamation + vbOKOnly
    End If
    
End Sub

