Sub DrawIsolatedFooting()
    ' Declare variables
    Dim acadApp As Object
    Dim acadDoc As Object
    Dim ms As Object
    Dim line1 As Object
    Dim p1 As Variant, p2 As Variant
    
    ' Initialize AutoCAD application
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application") ' Get running AutoCAD instance
    If acadApp Is Nothing Then
        Set acadApp = CreateObject("AutoCAD.Application") ' Open new instance if not running
    End If
    On Error GoTo 0
    
    acadApp.Visible = True ' Make sure AutoCAD is visible
    Set acadDoc = acadApp.ActiveDocument ' Get the active document
    Set ms = acadDoc.ModelSpace ' Get model space

    ' Define footing properties
    Dim side_cc As Double: side_cc = 75 ' Clear cover in mm
    Dim base_cc As Double: base_cc = 50 ' Base clear cover in mm
    Dim column_cc As Double: column_cc = 40 ' Column clear cover in mm
    
    ' Define dimensions (Assumed values, modify as needed)
    Dim l As Double: l = 1000 ' Footing length in mm
    Dim column_b As Double: column_b = 300 ' Column base width in mm
    Dim D As Double: D = 500 ' Depth of footing in mm
    
    ' Convert mm to meters (AutoCAD uses meters by default)
    l = l / 1000
    column_b = column_b / 1000
    D = D / 1000
    
    ' Define coordinates for footing plan
    Dim zz As Double: zz = 0
    Dim p1x(3) As Double, p1y(3) As Double
    Dim p2x(3) As Double, p2y(3) As Double
    
    p1x = Array(0, 0, l, l)
    p1y = Array(zz, zz + l, zz + l, zz)
    p2x = Array(0, l, l, 0)
    p2y = Array(zz + l, zz + l, zz, zz)
    
    ' Draw footing plan
    Dim i As Integer
    For i = 0 To 3
        p1 = Array(p1x(i), p1y(i), 0)
        p2 = Array(p2x(i), p2y(i), 0)
        Set line1 = ms.AddLine(p1, p2)
        line1.Layer = "Footing_Plan"
    Next i
    
    ' Define and draw column
    Dim col_p1x(3) As Double, col_p1y(3) As Double
    Dim col_p2x(3) As Double, col_p2y(3) As Double
    
    col_p1x = Array(l / 2 - column_b / 2, l / 2 - column_b / 2, l / 2 + column_b / 2, l / 2 + column_b / 2)
    col_p1y = Array(zz + l / 2 - column_b / 2, zz + l / 2 + column_b / 2, zz + l / 2 + column_b / 2, zz + l / 2 - column_b / 2)
    
    col_p2x = Array(l / 2 - column_b / 2, l / 2 + column_b / 2, l / 2 + column_b / 2, l / 2 - column_b / 2)
    col_p2y = Array(zz + l / 2 + column_b / 2, zz + l / 2 + column_b / 2, zz + l / 2 - column_b / 2, zz + l / 2 - column_b / 2)
    
    For i = 0 To 3
        p1 = Array(col_p1x(i), col_p1y(i), 0)
        p2 = Array(col_p2x(i), col_p2y(i), 0)
        Set line1 = ms.AddLine(p1, p2)
        line1.Layer = "Column"
    Next i

    ' Draw section
    Dim z As Double: z = -zz + 2
    Dim l_col As Double: l_col = 1
    
    Dim sec_p1x(4) As Double, sec_p1y(4) As Double
    Dim sec_p2x(4) As Double, sec_p2y(4) As Double
    
    sec_p1x = Array(0, 0, l / 2 + column_b / 2, 0, l)
    sec_p1y = Array(-(z + D), -z, -z, -z, -z)
    
    sec_p2x = Array(l, l / 2 - column_b / 2, l, 0, l)
    sec_p2y = Array(-(z + D), -z, -z, -(z + D), -(z + D))
    
    For i = 0 To 4
        p1 = Array(sec_p1x(i), sec_p1y(i), 0)
        p2 = Array(sec_p2x(i), sec_p2y(i), 0)
        Set line1 = ms.AddLine(p1, p2)
        line1.Layer = "Section_Line"
    Next i

    ' Draw bottom rebars
    Dim sp As Double: sp = (side_cc + 16) / 1000 ' Assuming dia_footing = 16 mm
    Dim No_rebar_final As Integer: No_rebar_final = 6 ' Example rebar count
    
    For i = 0 To No_rebar_final
        If i = 0 Then
            p1 = Array(sp, -(z + D) + base_cc / 1000, 0)
            p2 = Array(sp, -(z + D) + base_cc / 1000 + 0.1, 0)
        Else
            sp = sp + (l - 2 * (side_cc / 1000)) / (No_rebar_final - 1)
            p1 = Array(sp, -(z + D) + base_cc / 1000, 0)
            p2 = Array(sp, -(z + D) + base_cc / 1000 + 0.1, 0)
        End If
        
        Set line1 = ms.AddLine(p1, p2)
        line1.Layer = "Bottom_Rebar"
    Next i

    MsgBox "Isolated footing drawing completed!", vbInformation, "Success"
End Sub
