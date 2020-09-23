Attribute VB_Name = "RCTransparent"
Option Explicit

'set transparent form and all objects in form
Private Declare Function GetWindowLong _
    Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'--------------------------------------------------------------------------------
Private Declare Function SetWindowLong _
    Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'--------------------------------------------------------------------------------
Private Declare Function SetLayeredWindowAttributes _
    Lib "user32" (ByVal hwnd As Long, _
    ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'--------------------------------------------------------------------------------
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H8000F
Private Const LWA_ALPHA = &H2&

'another transparent function
'Create different types of regions declares
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
'POINTAPI type required for CreatePolygonRgn
Private Type POINTAPI
        X As Long
        Y As Long
End Type
'Sets the region
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'Combines the region
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
'Type of combine
Const RGN_XOR = 3

Public Sub Transparent(X As Form, LEVEL As Byte)
Call SetWindowLong _
    (X.hwnd, GWL_EXSTYLE, GetWindowLong(X.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
Call SetLayeredWindowAttributes(X.hwnd, 0, LEVEL, LWA_ALPHA)
End Sub

Public Sub MakeTransparent(TransForm As Form)
Dim ErrorTest As Double
    'In case there's an error, ignore it
    On Error Resume Next
    
    Dim Regn As Long
    Dim TmpRegn As Long
    Dim TmpControl As Control
    Dim LinePoints(4) As POINTAPI
    
    'Since the apis work with pixels, change the scalemode
    'To pixels
    TransForm.ScaleMode = 3
    
    'You have to have a borderless form, this just makes
    'sure it's borderless
    If TransForm.BorderStyle <> 0 Then MsgBox "Change the borderstyle to 0!", vbCritical, "ACK!": End
    
    'makes everything invisible
    Regn = CreateRectRgn(0, 0, 0, 0)
    
    'A loop to check every control in the form
    For Each TmpControl In TransForm
    
        'If the control is a line...
        If TypeOf TmpControl Is Line Then
            'Checks the slope
            If Abs((TmpControl.Y1 - TmpControl.Y2) / (TmpControl.X1 - TmpControl.X2)) > 1 Then
                'If it's more verticle than horizontal then
                'Set the points
                LinePoints(0).X = TmpControl.X1 - 1
                LinePoints(0).Y = TmpControl.Y1
                LinePoints(1).X = TmpControl.X2 - 1
                LinePoints(1).Y = TmpControl.Y2
                LinePoints(2).X = TmpControl.X2 + 1
                LinePoints(2).Y = TmpControl.Y2
                LinePoints(3).X = TmpControl.X1 + 1
                LinePoints(3).Y = TmpControl.Y1
            Else
                'If it's more horizontal than verticle then
                'Set the points
                LinePoints(0).X = TmpControl.X1
                LinePoints(0).Y = TmpControl.Y1 - 1
                LinePoints(1).X = TmpControl.X2
                LinePoints(1).Y = TmpControl.Y2 - 1
                LinePoints(2).X = TmpControl.X2
                LinePoints(2).Y = TmpControl.Y2 + 1
                LinePoints(3).X = TmpControl.X1
                LinePoints(3).Y = TmpControl.Y1 + 1
            End If
            'Creates the new polygon with the points
            TmpRegn = CreatePolygonRgn(LinePoints(0), 4, 1)
            
        'If the control is a shape...
        ElseIf TypeOf TmpControl Is Shape Then
            
            'An if that checks the type
            If TmpControl.Shape = 0 Then
            'It's a rectangle
                TmpRegn = CreateRectRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width, TmpControl.Top + TmpControl.Height)
            ElseIf TmpControl.Shape = 1 Then
            'It's a square
                If TmpControl.Width < TmpControl.Height Then
                    TmpRegn = CreateRectRgn(TmpControl.Left, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2, TmpControl.Left + TmpControl.Width, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width)
                Else
                    TmpRegn = CreateRectRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2, TmpControl.Top, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height, TmpControl.Top + TmpControl.Height)
                End If
            ElseIf TmpControl.Shape = 2 Then
            'It's an oval
                TmpRegn = CreateEllipticRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width + 0.5, TmpControl.Top + TmpControl.Height + 0.5)
            ElseIf TmpControl.Shape = 3 Then
            'It's a circle
                If TmpControl.Width < TmpControl.Height Then
                    TmpRegn = CreateEllipticRgn(TmpControl.Left, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2, TmpControl.Left + TmpControl.Width + 0.5, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width + 0.5)
                Else
                    TmpRegn = CreateEllipticRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2, TmpControl.Top, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height + 0.5, TmpControl.Top + TmpControl.Height + 0.5)
                End If
            ElseIf TmpControl.Shape = 4 Then
            'It's a rounded rectangle
                If TmpControl.Width > TmpControl.Height Then
                    TmpRegn = CreateRoundRectRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width + 1, TmpControl.Top + TmpControl.Height + 1, TmpControl.Height / 4, TmpControl.Height / 4)
                Else
                    TmpRegn = CreateRoundRectRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width + 1, TmpControl.Top + TmpControl.Height + 1, TmpControl.Width / 4, TmpControl.Width / 4)
                End If
            ElseIf TmpControl.Shape = 5 Then
            'It's a rounded square
                If TmpControl.Width > TmpControl.Height Then
                    TmpRegn = CreateRoundRectRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2, TmpControl.Top, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height + 1, TmpControl.Top + TmpControl.Height + 1, TmpControl.Height / 4, TmpControl.Height / 4)
                Else
                    TmpRegn = CreateRoundRectRgn(TmpControl.Left, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2, TmpControl.Left + TmpControl.Width + 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width + 1, TmpControl.Width / 4, TmpControl.Width / 4)
                End If
            End If
            
            'If the control is a shape with a transparent background
            If TmpControl.BackStyle = 0 Then
                
                'Combines the regions in memory and makes a new one
                CombineRgn Regn, Regn, TmpRegn, RGN_XOR
                
                If TmpControl.Shape = 0 Then
                'Rectangle
                    TmpRegn = CreateRectRgn(TmpControl.Left + 1, TmpControl.Top + 1, TmpControl.Left + TmpControl.Width - 1, TmpControl.Top + TmpControl.Height - 1)
                ElseIf TmpControl.Shape = 1 Then
                'Square
                    If TmpControl.Width < TmpControl.Height Then
                        TmpRegn = CreateRectRgn(TmpControl.Left + 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + 1, TmpControl.Left + TmpControl.Width - 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width - 1)
                    Else
                        TmpRegn = CreateRectRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + 1, TmpControl.Top + 1, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height - 1, TmpControl.Top + TmpControl.Height - 1)
                    End If
                ElseIf TmpControl.Shape = 2 Then
                'Oval
                    TmpRegn = CreateEllipticRgn(TmpControl.Left + 1, TmpControl.Top + 1, TmpControl.Left + TmpControl.Width - 0.5, TmpControl.Top + TmpControl.Height - 0.5)
                ElseIf TmpControl.Shape = 3 Then
                'Circle
                    If TmpControl.Width < TmpControl.Height Then
                        TmpRegn = CreateEllipticRgn(TmpControl.Left + 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + 1, TmpControl.Left + TmpControl.Width - 0.5, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width - 0.5)
                    Else
                        TmpRegn = CreateEllipticRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + 1, TmpControl.Top + 1, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height - 0.5, TmpControl.Top + TmpControl.Height - 0.5)
                    End If
                ElseIf TmpControl.Shape = 4 Then
                'Rounded rectangle
                    If TmpControl.Width > TmpControl.Height Then
                        TmpRegn = CreateRoundRectRgn(TmpControl.Left + 1, TmpControl.Top + 1, TmpControl.Left + TmpControl.Width, TmpControl.Top + TmpControl.Height, TmpControl.Height / 4, TmpControl.Height / 4)
                    Else
                        TmpRegn = CreateRoundRectRgn(TmpControl.Left + 1, TmpControl.Top + 1, TmpControl.Left + TmpControl.Width, TmpControl.Top + TmpControl.Height, TmpControl.Width / 4, TmpControl.Width / 4)
                    End If
                ElseIf TmpControl.Shape = 5 Then
                'Rounded square
                    If TmpControl.Width > TmpControl.Height Then
                        TmpRegn = CreateRoundRectRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + 1, TmpControl.Top + 1, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height, TmpControl.Top + TmpControl.Height, TmpControl.Height / 4, TmpControl.Height / 4)
                    Else
                        TmpRegn = CreateRoundRectRgn(TmpControl.Left + 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + 1, TmpControl.Left + TmpControl.Width, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width, TmpControl.Width / 4, TmpControl.Width / 4)
                    End If
                End If
            End If
        Else
                'Create a rectangular region with its parameters
                TmpRegn = CreateRectRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width, TmpControl.Top + TmpControl.Height)
            
        End If
            
            'Checks to make sure that the control has a width
            'or else you'll get some weird results
            ErrorTest = 0
            ErrorTest = TmpControl.Width
            If ErrorTest <> 0 Or TypeOf TmpControl Is Line Then
                'Combines the regions
                CombineRgn Regn, Regn, TmpRegn, RGN_XOR
            End If
        
    Next TmpControl
    
    'Make the regions
    SetWindowRgn TransForm.hwnd, Regn, True
    

End Sub
