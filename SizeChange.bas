Attribute VB_Name = "SizeChange"
'Dynamic control resizing module
'Scott Peterson
'IBPeterson@cox.net

'The purpose of this module is to dynamically resize all of the controls on a form
'when it is resized.

'In order to impliment this module you must call its initialize sub on form load
'and pass it the form name.

'In order for it to do the dynamic resizing you must call the sizechange sub from
'the form's resize event and pass it the form name.

'This builds a VDB or "virual database object" as I've called it for a control.
'I set the limit at 1000 which means that this will track up to 1000 controls total
'If you know how many conrols you have you can change the number.
Private controlinfo(1000) As vdb 'Creates references for the vdb objects.
Public lastcontrol As Integer 'Number representing the number of the last control added.

Public Sub initialize(formname As String)

Dim x As Integer 'Temporary variable used to cycle through forms
Dim y As Integer 'Temporary variable used to cycle through controls
Dim i As Integer 'Temporary variable used to track the number of the last control.

'This loop checks to see if the controls for a given form are loaded.
'This is done by checking all of the VDB objects stored form names for a match.
For x = 0 To Forms.Count - 1
    If Forms(x).Name = formname Then
        For i = 0 To lastcontrol - 1
            If Forms(x).Name = controlinfo(i).frm Then Exit Sub 'If it get's here then the form has already been added and we will not add it again.
        Next i
    End If
Next x

'If the form has not been added then we will go to the last position immediately
'after the last control that was added.
i = lastcontrol 'get last control number

'Here is where we will add each control from each form to the VDB list
For x = 0 To Forms.Count - 1
    If Forms(x).Name = formname Then
    For y = 0 To Forms(x).Controls.Count - 1
        'Yes this select case statement is ugly but diffrent controls have
        'different properties thus you must adjust accordingly
        'if a control type is not in this list you will need to add it in order
        'for this to work.  You will also have to change the sizechange sub.
        'Make sure you add the correct properties.  You may also need to change the
        'vdb class.
        Select Case TypeName(Forms(x).Controls(y))
            'This is the way a command button is handled
            Case Is = "CommandButton"
                Set controlinfo(i) = New vdb 'Creating the vdb object
                controlinfo(i).frm = Forms(x).Name 'moving form name to vdb object
                controlinfo(i).fw = Forms(x).Width 'moving forms width to vdb object
                controlinfo(i).fh = Forms(x).Height 'moving forms height to vdb object
                controlinfo(i).nm = Forms(x).Controls(y).Name 'moving control name to vdb object
                controlinfo(i).lft = Forms(x).Controls(y).Left 'moving controls left position to vdb object
                controlinfo(i).tp = Forms(x).Controls(y).Top 'moving controls top position to vdb object
                controlinfo(i).w = Forms(x).Controls(y).Width 'moving controls width to vdb object
                controlinfo(i).h = Forms(x).Controls(y).Height 'moving controls height to vdb object
                controlinfo(i).fntsize = Forms(x).Controls(y).FontSize 'moving controls font size to vdb object
                i = i + 1 'add one to the temporary last control variable
            'Other controls are handled the say way, though they may have different properties.
            Case Is = "ComboBox"
                Set controlinfo(i) = New vdb
                controlinfo(i).frm = Forms(x).Name
                controlinfo(i).fw = Forms(x).Width
                controlinfo(i).fh = Forms(x).Height
                controlinfo(i).nm = Forms(x).Controls(y).Name
                controlinfo(i).lft = Forms(x).Controls(y).Left
                controlinfo(i).tp = Forms(x).Controls(y).Top
                controlinfo(i).w = Forms(x).Controls(y).Width
                controlinfo(i).fntsize = Forms(x).Controls(y).FontSize
                i = i + 1
            Case Is = "Label"
                Set controlinfo(i) = New vdb
                controlinfo(i).frm = Forms(x).Name
                controlinfo(i).fw = Forms(x).Width
                controlinfo(i).fh = Forms(x).Height
                controlinfo(i).nm = Forms(x).Controls(y).Name
                controlinfo(i).lft = Forms(x).Controls(y).Left
                controlinfo(i).tp = Forms(x).Controls(y).Top
                controlinfo(i).w = Forms(x).Controls(y).Width
                controlinfo(i).h = Forms(x).Controls(y).Height
                controlinfo(i).fntsize = Forms(x).Controls(y).FontSize
                i = i + 1
            Case Is = "ListBox"
                Set controlinfo(i) = New vdb
                controlinfo(i).frm = Forms(x).Name
                controlinfo(i).fw = Forms(x).Width
                controlinfo(i).fh = Forms(x).Height
                controlinfo(i).nm = Forms(x).Controls(y).Name
                controlinfo(i).lft = Forms(x).Controls(y).Left
                controlinfo(i).tp = Forms(x).Controls(y).Top
                controlinfo(i).w = Forms(x).Controls(y).Width
                controlinfo(i).h = Forms(x).Controls(y).Height
                controlinfo(i).fntsize = Forms(x).Controls(y).FontSize
                i = i + 1
            Case Is = "TextBox"
                Set controlinfo(i) = New vdb
                controlinfo(i).frm = Forms(x).Name
                controlinfo(i).fw = Forms(x).Width
                controlinfo(i).fh = Forms(x).Height
                controlinfo(i).nm = Forms(x).Controls(y).Name
                controlinfo(i).lft = Forms(x).Controls(y).Left
                controlinfo(i).tp = Forms(x).Controls(y).Top
                controlinfo(i).w = Forms(x).Controls(y).Width
                controlinfo(i).h = Forms(x).Controls(y).Height
                controlinfo(i).fntsize = Forms(x).Controls(y).FontSize
                i = i + 1
            Case Is = "CheckBox"
                Set controlinfo(i) = New vdb
                controlinfo(i).frm = Forms(x).Name
                controlinfo(i).fw = Forms(x).Width
                controlinfo(i).fh = Forms(x).Height
                controlinfo(i).nm = Forms(x).Controls(y).Name
                controlinfo(i).w = Forms(x).Controls(y).Width
                controlinfo(i).h = Forms(x).Controls(y).Height
                controlinfo(i).fntsize = Forms(x).Controls(y).FontSize
                controlinfo(i).lft = Forms(x).Controls(y).Left
                controlinfo(i).tp = Forms(x).Controls(y).Top
                i = i + 1
                
            Case Is = "Calendar"
                Set controlinfo(i) = New vdb
                controlinfo(i).frm = Forms(x).Name
                controlinfo(i).fw = Forms(x).Width
                controlinfo(i).fh = Forms(x).Height
                controlinfo(i).nm = Forms(x).Controls(y).Name
                controlinfo(i).lft = Forms(x).Controls(y).Left
                controlinfo(i).tp = Forms(x).Controls(y).Top
                controlinfo(i).w = Forms(x).Controls(y).Width
                controlinfo(i).h = Forms(x).Controls(y).Height
                i = i + 1
            Case Is = "Line"
                Set controlinfo(i) = New vdb
                controlinfo(i).frm = Forms(x).Name
                controlinfo(i).fw = Forms(x).Width
                controlinfo(i).fh = Forms(x).Height
                controlinfo(i).nm = Forms(x).Controls(y).Name
                controlinfo(i).x1 = Forms(x).Controls(y).x1
                controlinfo(i).x2 = Forms(x).Controls(y).x2
                controlinfo(i).y1 = Forms(x).Controls(y).y1
                controlinfo(i).y2 = Forms(x).Controls(y).y2
                i = i + 1
            Case Is = "OptionButton"
                Set controlinfo(i) = New vdb
                controlinfo(i).frm = Forms(x).Name
                controlinfo(i).fw = Forms(x).Width
                controlinfo(i).fh = Forms(x).Height
                controlinfo(i).nm = Forms(x).Controls(y).Name
                controlinfo(i).lft = Forms(x).Controls(y).Left
                controlinfo(i).tp = Forms(x).Controls(y).Top
                controlinfo(i).w = Forms(x).Controls(y).Width
                controlinfo(i).h = Forms(x).Controls(y).Height
                controlinfo(i).fntsize = Forms(x).Controls(y).FontSize
                i = i + 1
            Case Is = "Frame"
                Set controlinfo(i) = New vdb
                controlinfo(i).frm = Forms(x).Name
                controlinfo(i).fw = Forms(x).Width
                controlinfo(i).fh = Forms(x).Height
                controlinfo(i).nm = Forms(x).Controls(y).Name
                controlinfo(i).lft = Forms(x).Controls(y).Left
                controlinfo(i).tp = Forms(x).Controls(y).Top
                controlinfo(i).w = Forms(x).Controls(y).Width
                controlinfo(i).h = Forms(x).Controls(y).Height
                controlinfo(i).fntsize = Forms(x).Controls(y).FontSize
                i = i + 1
            Case Is = "DriveListBox"
                Set controlinfo(i) = New vdb
                controlinfo(i).frm = Forms(x).Name
                controlinfo(i).fw = Forms(x).Width
                controlinfo(i).fh = Forms(x).Height
                controlinfo(i).nm = Forms(x).Controls(y).Name
                controlinfo(i).lft = Forms(x).Controls(y).Left
                controlinfo(i).tp = Forms(x).Controls(y).Top
                controlinfo(i).w = Forms(x).Controls(y).Width
                controlinfo(i).fntsize = Forms(x).Controls(y).FontSize
                i = i + 1
            Case Is = "DirListBox"
                Set controlinfo(i) = New vdb
                controlinfo(i).frm = Forms(x).Name
                controlinfo(i).fw = Forms(x).Width
                controlinfo(i).fh = Forms(x).Height
                controlinfo(i).nm = Forms(x).Controls(y).Name
                controlinfo(i).lft = Forms(x).Controls(y).Left
                controlinfo(i).tp = Forms(x).Controls(y).Top
                controlinfo(i).w = Forms(x).Controls(y).Width
                controlinfo(i).h = Forms(x).Controls(y).Height
                controlinfo(i).fntsize = Forms(x).Controls(y).FontSize
                i = i + 1
            Case Is = "FileListBox"
                Set controlinfo(i) = New vdb
                controlinfo(i).frm = Forms(x).Name
                controlinfo(i).fw = Forms(x).Width
                controlinfo(i).fh = Forms(x).Height
                controlinfo(i).nm = Forms(x).Controls(y).Name
                controlinfo(i).lft = Forms(x).Controls(y).Left
                controlinfo(i).tp = Forms(x).Controls(y).Top
                controlinfo(i).w = Forms(x).Controls(y).Width
                controlinfo(i).h = Forms(x).Controls(y).Height
                controlinfo(i).fntsize = Forms(x).Controls(y).FontSize
                i = i + 1
            Case Is = "MediaPlayer"
                Set controlinfo(i) = New vdb
                controlinfo(i).frm = Forms(x).Name
                controlinfo(i).fw = Forms(x).Width
                controlinfo(i).fh = Forms(x).Height
                controlinfo(i).nm = Forms(x).Controls(y).Name
                controlinfo(i).lft = Forms(x).Controls(y).Left
                controlinfo(i).tp = Forms(x).Controls(y).Top
                controlinfo(i).w = Forms(x).Controls(y).Width
                controlinfo(i).h = Forms(x).Controls(y).Height
                i = i + 1
            Case Is = "Image"
                Set controlinfo(i) = New vdb
                controlinfo(i).frm = Forms(x).Name
                controlinfo(i).fw = Forms(x).Width
                controlinfo(i).fh = Forms(x).Height
                controlinfo(i).nm = Forms(x).Controls(y).Name
                controlinfo(i).lft = Forms(x).Controls(y).Left
                controlinfo(i).tp = Forms(x).Controls(y).Top
                controlinfo(i).w = Forms(x).Controls(y).Width
                controlinfo(i).h = Forms(x).Controls(y).Height
                i = i + 1
            Case Is = "HScrollBar"
                Set controlinfo(i) = New vdb
                controlinfo(i).frm = Forms(x).Name
                controlinfo(i).fw = Forms(x).Width
                controlinfo(i).fh = Forms(x).Height
                controlinfo(i).nm = Forms(x).Controls(y).Name
                controlinfo(i).lft = Forms(x).Controls(y).Left
                controlinfo(i).tp = Forms(x).Controls(y).Top
                controlinfo(i).w = Forms(x).Controls(y).Width
                controlinfo(i).h = Forms(x).Controls(y).Height
                i = i + 1
            Case Is = "VScrollBar"
                Set controlinfo(i) = New vdb
                controlinfo(i).frm = Forms(x).Name
                controlinfo(i).fw = Forms(x).Width
                controlinfo(i).fh = Forms(x).Height
                controlinfo(i).nm = Forms(x).Controls(y).Name
                controlinfo(i).lft = Forms(x).Controls(y).Left
                controlinfo(i).tp = Forms(x).Controls(y).Top
                controlinfo(i).w = Forms(x).Controls(y).Width
                controlinfo(i).h = Forms(x).Controls(y).Height
                i = i + 1
            Case Else
        End Select
    Next y
    End If
Next x

lastcontrol = i 'This moves the temporary variable back to last control.
End Sub
Public Sub SizeChange(formname As String)
'This is a constant value representing form width it was required for a MDI application
'Const fw = 4725
'This is a constant value representing form height it was required for a MDI application
'Const fh = 5430

Dim x As Integer 'Temporary variable used to cycle through forms
Dim y As Integer 'Temporary variable used to cycle through controls
Dim i As Integer 'Temporary variable used to track the number of the last control
Dim wp As Double 'Temporary variable used to cycle track the percent change in form width
Dim hp As Double 'Temporary variable used to cycle track the percent change in form height

'This is where everything will get resized.
For x = 0 To Forms.Count - 1
    If Forms(x).Name = formname Then 'First we make sure the form name that is passed matches the one were checking in the forms collection
    For y = 0 To Forms(x).Controls.Count - 1
        For i = 0 To lastcontrol - 1
            If Forms(x).Name = controlinfo(i).frm Then 'Now we make sure the form name in the forms collection matches the one in the VDB list
                If Forms(x).Controls(y).Name = controlinfo(i).nm Then 'Now we match up the control names
                    wp = Forms(x).Width / controlinfo(i).fw 'here we get the percent change in form width
                    hp = Forms(x).Height / controlinfo(i).fh 'here we get the percent change in form height
                    'Yes this select case statement is also ugly but diffrent controls have
                    'different properties thus you must adjust accordingly
                    'if a control type is not in this list you will need to add it in order
                    'for this to work.  You will also have to change the initialize sub.
                    'Make sure you add the correct properties.  You may also need to change the
                    'vdb class.
                    Select Case TypeName(Forms(x).Controls(y))
                        'This is the way a command button is handled
                        Case Is = "CommandButton"
                            Forms(x).Controls(y).Left = Int(controlinfo(i).lft * wp) 'We adjust the size according to what we previously stored times the width percent
                            Forms(x).Controls(y).Top = Int(controlinfo(i).tp * hp) 'We adjust the size according to what we previously stored times the height percent
                            Forms(x).Controls(y).Width = Int(controlinfo(i).w * wp) 'We adjust the size according to what we previously stored times the width percent
                            Forms(x).Controls(y).Height = Int(controlinfo(i).h * hp) 'We adjust the size according to what we previously stored times the height percent
                            'Font size is a special case where the font size cannot be allowed to be set to 0.
                            If Int(controlinfo(i).fntsize * wp) > 0 And Int(controlinfo(i).fntsize * hp) > 0 Then
                                'This is where we take the percentage size of width and height and base the new font size on the smallest percentage
                                If hp > wp Then
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * wp)
                                Else
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * hp)
                                End If
                            End If
                        'Other controls are handled the say way, though they may have different properties.
                        Case Is = "ComboBox"
                            Forms(x).Controls(y).Left = Int(controlinfo(i).lft * wp)
                            Forms(x).Controls(y).Top = Int(controlinfo(i).tp * hp)
                            Forms(x).Controls(y).Width = Int(controlinfo(i).w * wp)
                            If Int(controlinfo(i).fntsize * wp) > 0 And Int(controlinfo(i).fntsize * hp) > 0 Then
                                If hp > wp Then
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * wp)
                                Else
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * hp)
                                End If
                            End If
                        Case Is = "Label"
                            Forms(x).Controls(y).Left = Int(controlinfo(i).lft * wp)
                            Forms(x).Controls(y).Top = Int(controlinfo(i).tp * hp)
                            Forms(x).Controls(y).Width = Int(controlinfo(i).w * wp)
                            Forms(x).Controls(y).Height = Int(controlinfo(i).h * hp)
                            If Int(controlinfo(i).fntsize * wp) > 0 And Int(controlinfo(i).fntsize * hp) > 0 Then
                                If hp > wp Then
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * wp)
                                Else
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * hp)
                                End If
                            End If
                        Case Is = "ListBox"
                            Forms(x).Controls(y).Left = Int(controlinfo(i).lft * wp)
                            Forms(x).Controls(y).Top = Int(controlinfo(i).tp * hp)
                            Forms(x).Controls(y).Width = Int(controlinfo(i).w * wp)
                            Forms(x).Controls(y).Height = Int(controlinfo(i).h * hp)
                            If Int(controlinfo(i).fntsize * wp) > 0 And Int(controlinfo(i).fntsize * hp) > 0 Then
                                If hp > wp Then
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * wp)
                                Else
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * hp)
                                End If
                            End If
                        Case Is = "TextBox"
                            Forms(x).Controls(y).Left = Int(controlinfo(i).lft * wp)
                            Forms(x).Controls(y).Top = Int(controlinfo(i).tp * hp)
                            Forms(x).Controls(y).Width = Int(controlinfo(i).w * wp)
                            Forms(x).Controls(y).Height = Int(controlinfo(i).h * hp)
                            If Int(hp) > 0 And Int(wp) > 0 Then
                                If hp > wp Then
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * wp)
                                Else
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * hp)
                                End If
                            End If
                        Case Is = "Calendar"
                            Forms(x).Controls(y).Width = Int(controlinfo(i).w * wp)
                            Forms(x).Controls(y).Height = Int(controlinfo(i).h * hp)
                            Forms(x).Controls(y).Height = Int(controlinfo(i).h * hp)
                        Case Is = "Line"
                            Forms(x).Controls(y).x1 = Int(controlinfo(i).x1 * wp)
                            Forms(x).Controls(y).y1 = Int(controlinfo(i).y1 * hp)
                            Forms(x).Controls(y).x2 = Int(controlinfo(i).x2 * wp)
                            Forms(x).Controls(y).y2 = Int(controlinfo(i).y2 * hp)
                        Case Is = "CheckBox"
                            Forms(x).Controls(y).Left = Int(controlinfo(i).lft * wp)
                            Forms(x).Controls(y).Top = Int(controlinfo(i).tp * hp)
                            Forms(x).Controls(y).Width = Int(controlinfo(i).w * wp)
                            Forms(x).Controls(y).Height = Int(controlinfo(i).h * hp)
                            If Int(controlinfo(i).fntsize * wp) > 0 And Int(controlinfo(i).fntsize * hp) > 0 Then
                                If hp > wp Then
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * wp)
                                Else
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * hp)
                                End If
                            End If
                        Case Is = "OptionButton"
                            Forms(x).Controls(y).Left = Int(controlinfo(i).lft * wp)
                            Forms(x).Controls(y).Top = Int(controlinfo(i).tp * hp)
                            Forms(x).Controls(y).Width = Int(controlinfo(i).w * wp)
                            Forms(x).Controls(y).Height = Int(controlinfo(i).h * hp)
                            If Int(hp) > 0 And Int(wp) > 0 Then
                                If hp > wp Then
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * wp)
                                Else
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * hp)
                                End If
                            End If
                        Case Is = "Frame"
                            Forms(x).Controls(y).Left = Int(controlinfo(i).lft * wp)
                            Forms(x).Controls(y).Top = Int(controlinfo(i).tp * hp)
                            Forms(x).Controls(y).Width = Int(controlinfo(i).w * wp)
                            Forms(x).Controls(y).Height = Int(controlinfo(i).h * hp)
                            If Int(controlinfo(i).fntsize * wp) > 0 And Int(controlinfo(i).fntsize * hp) > 0 Then
                                If hp > wp Then
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * wp)
                                Else
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * hp)
                                End If
                            End If
                        Case Is = "DriveListBox"
                            Forms(x).Controls(y).Left = Int(controlinfo(i).lft * wp)
                            Forms(x).Controls(y).Top = Int(controlinfo(i).tp * hp)
                            Forms(x).Controls(y).Width = Int(controlinfo(i).w * wp)
                            If Int(controlinfo(i).fntsize * wp) > 0 And Int(controlinfo(i).fntsize * hp) > 0 Then
                                If hp > wp Then
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * wp)
                                Else
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * hp)
                                End If
                            End If
                        Case Is = "DirListBox"
                            Forms(x).Controls(y).Left = Int(controlinfo(i).lft * wp)
                            Forms(x).Controls(y).Top = Int(controlinfo(i).tp * hp)
                            Forms(x).Controls(y).Width = Int(controlinfo(i).w * wp)
                            Forms(x).Controls(y).Height = Int(controlinfo(i).h * hp)
                            If Int(controlinfo(i).fntsize * wp) > 0 And Int(controlinfo(i).fntsize * hp) > 0 Then
                                If hp > wp Then
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * wp)
                                Else
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * hp)
                                End If
                            End If
                        Case Is = "FileListBox"
                            Forms(x).Controls(y).Left = Int(controlinfo(i).lft * wp)
                            Forms(x).Controls(y).Top = Int(controlinfo(i).tp * hp)
                            Forms(x).Controls(y).Width = Int(controlinfo(i).w * wp)
                            Forms(x).Controls(y).Height = Int(controlinfo(i).h * hp)
                            If Int(controlinfo(i).fntsize * wp) > 0 And Int(controlinfo(i).fntsize * hp) > 0 Then
                                If hp > wp Then
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * wp)
                                Else
                                    Forms(x).Controls(y).FontSize = Int(controlinfo(i).fntsize * hp)
                                End If
                            End If
                        Case Is = "MediaPlayer"
                            Forms(x).Controls(y).Left = Int(controlinfo(i).lft * wp)
                            Forms(x).Controls(y).Top = Int(controlinfo(i).tp * hp)
                            Forms(x).Controls(y).Width = Int(controlinfo(i).w * wp)
                            Forms(x).Controls(y).Height = Int(controlinfo(i).h * hp)
                        Case Is = "Image"
                            Forms(x).Controls(y).Left = Int(controlinfo(i).lft * wp)
                            Forms(x).Controls(y).Top = Int(controlinfo(i).tp * hp)
                            Forms(x).Controls(y).Width = Int(controlinfo(i).w * wp)
                            Forms(x).Controls(y).Height = Int(controlinfo(i).h * hp)
                        Case Is = "HScrollBar"
                            Forms(x).Controls(y).Left = Int(controlinfo(i).lft * wp)
                            Forms(x).Controls(y).Top = Int(controlinfo(i).tp * hp)
                            Forms(x).Controls(y).Width = Int(controlinfo(i).w * wp)
                            Forms(x).Controls(y).Height = Int(controlinfo(i).h * hp)
                        Case Is = "VScrollBar"
                            Forms(x).Controls(y).Left = Int(controlinfo(i).lft * wp)
                            Forms(x).Controls(y).Top = Int(controlinfo(i).tp * hp)
                            Forms(x).Controls(y).Width = Int(controlinfo(i).w * wp)
                            Forms(x).Controls(y).Height = Int(controlinfo(i).h * hp)
                        Case Else
                    End Select
                End If
            End If
        Next i
    Next y
    End If
Next x
End Sub
