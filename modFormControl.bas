Attribute VB_Name = "modFormControl"
Private List() As Control
Private curr_obj As Object
Private iHeight As Integer
Private iWidth As Integer
Private x_size As Double
Private y_size As Double
'*****************************************************************************************
'                           LICENSE INFORMATION
'*****************************************************************************************
'   FormControl Version 2.0
'   Code module for resizing a form based on screen size, then resizing the
'   controls based on the forms size
'
'   Copyright (C) 2007
'   Richard L. McCutchen
'   Email: richard@psychocoder.net
'   Created: AUG99
'
'   This program is free software: you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or
'   (at your option) any later version.
'
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with this program.  If not, see <http://www.gnu.org/licenses/>.
'*****************************************************************************************

Private Type Control
    '''Index As Integer
    Index As String
    Name As String
    Left As Integer
    Top As Integer
    width As Integer
    height As Integer
    fontSize As Integer
End Type

Public Sub ResizeControls(frm As Form)

Dim i As Integer
'   Get ratio of initial form size to current form size
x_size = frm.height / iHeight
y_size = frm.width / iWidth

'Loop though all the objects on the form
'Based on the upper bound of the # of controls
For i = 0 To UBound(List)
    On Error GoTo ErrorCode
    'Grad each control individually
    For Each curr_obj In frm
        On Error Resume Next
        'Check to make sure its the right control
        If curr_obj.TabIndex = List(i).Index Then
            'Then resize the control
             With curr_obj
                .Left = List(i).Left * y_size
                .width = List(i).width * y_size
                .height = List(i).height * x_size
                .Top = List(i).Top * x_size
             End With
        End If
DebugCode:
        If Err.Number = 438 Then Err.Clear
    'Get the next control
    Next curr_obj
ErrorCode:
    If Err.Number = 438 Then Err.Clear

Next i
End Sub

Public Function SetFontSize() As Integer
    'Make sure x_size is greater than 0
    If Int(x_size) > 0 Then
    'Set the font size
        SetFontSize = Int(x_size * 8)
    End If
End Function

Public Sub GetLocation(frm As Form)
Dim i As Integer
'   Load the current positions of each object into a user defined type array.
'   This information will be used to rescale them in the Resize function.
On Error Resume Next
i = 0
'Loop through each control
For Each curr_obj In frm
    
'Resize the Array by 1, and preserve
'the original objects in the array
    ReDim Preserve List(i)
    With List(i)
        .Name = curr_obj.Name()
        .Index = curr_obj.Index
        .Left = curr_obj.Left
        .Top = curr_obj.Top
        .width = curr_obj.width
        .height = curr_obj.height
    End With
    i = i + 1
Next curr_obj
    
'   This is what the object sizes will be compared to on rescaling.
    iHeight = frm.height
    iWidth = frm.width
End Sub

Public Sub CenterForm(frm As Form)
    frm.Move (Screen.width - frm.width) \ 2, (Screen.height - frm.height) \ 2
End Sub

Public Sub ResizeForm(frm As Form)
    'Set the forms height
    frm.height = Screen.height / 2
    'Set the forms width
    frm.width = Screen.width / 2
    'Resize all of the controls
    'based on the forms new size
    ResizeControls frm
End Sub


