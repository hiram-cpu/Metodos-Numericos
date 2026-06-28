VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10452
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal x As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub aprox_i_Change()
    validar_new
End Sub

Private Sub aprox_i_m_Click()
    aprox_i.Visible = True
End Sub

Private Sub aprox_i_m_new_m_Click()
    aprox_i_new_m.Visible = True
End Sub

Private Sub aprox_i_new_m_Change()
    validar_new_m
End Sub

Private Sub aprox_i_pm_Click()
    aprox_i.Value = ""
    aprox_i.Visible = False
End Sub

Private Sub aprox_i_pm_new_m_Click()
    aprox_i_new_m.Value = ""
    aprox_i_new_m.Visible = False
End Sub

Private Sub bisección_Click()
    siguiente2.Enabled = True
End Sub

Private Sub falsi_Click()
    siguiente2.Enabled = True
End Sub

Private Sub inferior_bis_Change()
    validar_bis
End Sub

Private Sub inferior_Change()
    validar
End Sub

Private Sub inferior_new_Change()
    validar_new
End Sub

Private Sub inferior_new_m_Change()
    validar_new_m
End Sub

Private Sub inferior_sec_Change()
    validar_sec
End Sub

Private Sub iter_bis_Change()
    validar_bis
End Sub

Private Sub iter_Change()
    validar
End Sub

Private Sub iter_new_Change()
    validar_new
End Sub

Private Sub iter_new_m_Change()
    validar_new_m
End Sub

Private Sub iter_sec_Change()
    validar_sec
End Sub

Private Sub label_iter_Click()

End Sub

Private Sub label_iter_new_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label25_Click()

End Sub

Private Sub newton_Click()
    siguiente2.Enabled = True
End Sub

Private Sub newton_m_Click()
    siguiente2.Enabled = True
End Sub

Private Sub opcion_iter_bis_Click()
    label_tol_bis.Visible = False
    label_iter_bis.Visible = True
    tol_bis.Visible = False
    iter_bis.Visible = True
    tol_bis.Value = ""
End Sub

Private Sub opcion_iter_new_Click()
    label_tol_new.Visible = False
    label_iter_new.Visible = True
    tol_new.Visible = False
    iter_new.Visible = True
    tol_new.Value = ""
End Sub

Private Sub opcion_iter_new_m_Click()
    label_tol_new_m.Visible = False
    label_iter_new_m.Visible = True
    tol_new_m.Visible = False
    iter_new_m.Visible = True
    tol_new_m.Value = ""
End Sub

Private Sub opcion_iter_sec_Click()
    label_tol_sec.Visible = False
    label_iter_sec.Visible = True
    tol_sec.Visible = False
    iter_sec.Visible = True
    tol_sec.Value = ""
End Sub

Private Sub opcion_iteraciones_Click()
    label_tol.Visible = False
    label_iter.Visible = True
    tol.Visible = False
    iter.Visible = True
    tol.Value = ""
End Sub

Private Sub opcion_tol_bis_Click()
    label_tol_bis.Visible = True
    label_iter_bis.Visible = False
    tol_bis.Visible = True
    iter_bis.Visible = False
    iter_bis.Value = ""
End Sub

Private Sub opcion_tol_new_Click()
    label_tol_new.Visible = True
    label_iter_new.Visible = False
    tol_new.Visible = True
    iter_new.Visible = False
    iter_new.Value = ""
End Sub

Private Sub opcion_tol_new_m_Click()
    label_tol_new_m.Visible = True
    label_iter_new_m.Visible = False
    tol_new_m.Visible = True
    iter_new_m.Visible = False
    iter_new_m.Value = ""
End Sub

Private Sub opcion_tol_sec_Click()
    label_tol_sec.Visible = True
    label_iter_sec.Visible = False
    tol_sec.Visible = True
    iter_sec.Visible = False
    iter_sec.Value = ""
End Sub

Private Sub opcion_tolerancia_Click()
    label_tol.Visible = True
    label_iter.Visible = False
    tol.Visible = True
    iter.Visible = False
    iter.Value = ""
End Sub

Private Sub OptionButton6_Click()

End Sub

Private Sub regresar_bis_Click()
    UserForm1.MultiPage1.Value = 0
    bisección.Value = False
    siguiente.Enabled = False
End Sub

Private Sub regresar_falsi_Click()
    UserForm1.MultiPage1.Value = 0
    falsi.Value = False
    siguiente.Enabled = False
    End Sub

Private Sub regresar_new_Click()
    UserForm1.MultiPage1.Value = 0
    newton.Value = False
    siguiente.Enabled = False
End Sub

Private Sub regresar_new_m_Click()
    UserForm1.MultiPage1.Value = 0
    newton_m.Value = False
    siguiente.Enabled = False
    
End Sub

Private Sub regresar_sec_Click()
    UserForm1.MultiPage1.Value = 0
    secante.Value = False
    siguiente.Enabled = False
End Sub

Private Sub salir_bis_Click()
    Unload Me
End Sub

Private Sub salir_Click()
    Unload Me 'Finaliza la macro
End Sub

Private Sub salir_new_Click()
    Unload Me
End Sub

Private Sub salir_new_m_Click()
    Unload Me
End Sub

Private Sub salir_sec_Click()
    Unload Me
End Sub

Private Sub salir2_Click()
    Unload Me
End Sub

Private Sub secante_Click()
    siguiente2.Enabled = True
End Sub

Private Sub siguiente_bis_Click()
    Dim a As Double
    Dim b As Double
    Dim Tolerancia As Double
    Dim n_iteración As Double
    Dim c As Double
    Dim Iteración As Long
    Dim i As Double
    Dim criterio As Integer
        
    Sheets("Hoja1").Select
    Cells.Clear
    
    If fun1(inferior_bis.Value) * fun1(superior_bis.Value) > 0 Then 'Se verifica la existencia de una raíz de la función en el intervalo creado por a y b
        MsgBox "No se halló raíz alguna en el intervalo creado por a = " & a & " y b = " & b & ", ingrese nuevos valores para ambos límites"
        inferior_bis.Value = ""
        superior_bis.Value = ""
        tol_bis.Value = ""
        iter_bis.Value = ""
        inferior_bis.SetFocus
        Exit Sub
    End If
    
    Range("A1") = "Límite Inferior"
    Range("A2") = "Límite Superior"
    Range("A3") = "Criterio"
    Range("B1") = inferior_bis.Value
    Range("B2") = superior_bis.Value
    
    a = inferior_bis.Value
    b = superior_bis.Value
    c = (a + b) / 2
    Iteración = 0
    
    Range("A6:E6") = Array("Iteración", "a", "b", "c", "|f(c)|")
    
    If opcion_tol_bis.Value = True Then
        Range("B3") = "Margen de Tolerancia"
        Range("A4") = "Tolerancia"
        Range("B4") = tol_bis.Value
        criterio = 1
    Else
        Range("B3") = "Máximo de iteraciones"
        Range("A4") = "Iteraciones"
        Range("B4") = iter_bis.Value
        criterio = 2
    End If
    
    Select Case criterio
        Case 1
        Do While Abs(fun1(c)) > Val(tol_bis.Value)
            c = (a + b) / 2
            Iteración = Iteración + 1
            Cells(Iteración + 6, 1) = Iteración
            Cells(Iteración + 6, 2) = a
            Cells(Iteración + 6, 3) = b
            Cells(Iteración + 6, 4) = c
            Cells(Iteración + 6, 5) = Abs(fun1(c))
            If fun1(a) * fun1(c) > 0 Then
                a = c
                b = b
            Else
                a = a
                b = c
            End If
        Loop
        
        Case 2
        For i = 1 To Val(iter_bis.Value)
            c = (a + b) / 2
            Iteración = Iteración + 1
            Cells(Iteración + 6, 1) = Iteración
            Cells(Iteración + 6, 2) = a
            Cells(Iteración + 6, 3) = b
            Cells(Iteración + 6, 4) = c
            Cells(Iteración + 6, 5) = Abs(fun1(c))
            If fun1(a) * fun1(c) > 0 Then
                a = c
                b = b
            Else
                a = a
                b = c
            End If
        Next i
    End Select
        
    'Se resalta la celda que contiene el resultado de la última iteración
    Range("D" & Iteración + 6).Interior.Color = RGB(117, 219, 255)
    
    'Se da formato de tabla a las tablas generadas
    Range("A6:E" & Iteración + 6).Borders.LineStyle = xlContinuous
    Range("A1:B4").Borders.LineStyle = xlContinuous
    
    'Se ajusta el tamaño de las celdas al contenido de estas
    Cells.Select
    Cells.EntireColumn.AutoFit
        
End Sub


Private Sub siguiente_Click()
    Dim a As Double
    Dim b As Double
    Dim Tolerancia As Double
    Dim n_iteración As Double
    Dim c As Double
    Dim Iteración As Long
    Dim i As Double
    Dim criterio As Integer
    
    Sheets("Hoja1").Select
    Cells.Clear
        
    If fun1(inferior.Value) * fun1(superior.Value) > 0 Then 'Se verifica la existencia de una raíz de la función en el intervalo creado por a y b
        MsgBox "No se halló raíz alguna en el intervalo creado por a = " & a & " y b = " & b & ", ingrese nuevos valores para ambos límites"
        inferior.Value = ""
        superior.Value = ""
        tol.Value = ""
        iter.Value = ""
        inferior.SetFocus
        Exit Sub
    End If
    
    Range("A1") = "Límite Inferior"
    Range("A2") = "Límite Superior"
    Range("A3") = "Criterio"
    Range("B1") = inferior.Value
    Range("B2") = superior.Value
    
    a = inferior.Value
    b = superior.Value
    c = b - fun1(b) * (a - b) / (fun1(a) - fun1(b))
    Iteración = 0
    
    Range("A6:E6") = Array("Iteración", "a", "b", "c", "|f(c)|")
    
    If opcion_tolerancia.Value = True Then
        Range("B3") = "Margen de Tolerancia"
        Range("A4") = "Tolerancia"
        Range("B4") = tol.Value
        criterio = 1
    Else
        Range("B3") = "Máximo de iteraciones"
        Range("A4") = "Iteraciones"
        Range("B4") = iter.Value
        criterio = 2
    End If
    
    Select Case criterio
        Case 1
        Do While Abs(fun1(c)) > Val(tol.Value)
            c = b - fun1(b) * (a - b) / (fun1(a) - fun1(b))
            Iteración = Iteración + 1
            Cells(Iteración + 6, 1) = Iteración
            Cells(Iteración + 6, 2) = a
            Cells(Iteración + 6, 3) = b
            Cells(Iteración + 6, 4) = c
            Cells(Iteración + 6, 5) = Abs(fun1(c))
            If fun1(a) * fun1(c) > 0 Then
                a = c
                b = b
            Else
                a = a
                b = c
            End If
        Loop
        
        Case 2
        For i = 1 To Val(iter.Value)
        c = b - fun1(b) * (a - b) / (fun1(a) - fun1(b))
        Iteración = Iteración + 1
        Cells(Iteración + 6, 1) = Iteración
        Cells(Iteración + 6, 2) = a
        Cells(Iteración + 6, 3) = b
        Cells(Iteración + 6, 4) = c
        Cells(Iteración + 6, 5) = Abs(fun1(c))
        If fun1(a) * fun1(c) > 0 Then
            a = c
            b = b
        Else
            a = a
            b = c
        End If
        Next i
    End Select
        
    'Se resalta la celda que contiene el resultado de la última iteración
    Range("D" & Iteración + 6).Interior.Color = RGB(117, 219, 255)
    
    'Se da formato de tabla a las tablas generadas
    Range("A6:E" & Iteración + 6).Borders.LineStyle = xlContinuous
    Range("A1:B4").Borders.LineStyle = xlContinuous
    
    'Se ajusta el tamaño de las celdas al contenido de estas
    Cells.Select
    Cells.EntireColumn.AutoFit
        
End Sub

Private Sub siguiente_new_Click()
    Dim a As Double
    Dim b As Double
    Dim p_n As Double
    Dim Tolerancia As Double
    Dim Iteración As Long
    Dim i As Double
    Dim criterio As Integer
    
    Sheets("Hoja1").Select
    Cells.Clear
    
    ' Se asignan las variables iniciales usando Val() para evitar errores de texto
    a = Val(inferior_new.Value)
    b = Val(superior_new.Value)
    
    If fun1(a) * fun1(b) > 0 Then ' Se verifica la existencia de una raíz
        MsgBox "No se halló raíz alguna en el intervalo creado por a = " & a & " y b = " & b & ", ingrese nuevos valores para ambos límites"
        inferior_new.Value = ""
        superior_new.Value = ""
        tol_new.Value = ""
        iter_new.Value = ""
        inferior_new.SetFocus
        Exit Sub
    End If
    
    Range("A1") = "Límite Inferior"
    Range("A2") = "Límite Superior"
    Range("A3") = "Criterio"
    Range("A5") = "Aprox. Inicial"
    Range("B1") = a
    Range("B2") = b
    
    ' 1. Definir la aproximación inicial real (Iteración 0)
    If aprox_i_pm.Value = True Then
        p_n = (a + b) / 2
    Else
        p_n = Val(aprox_i.Value)
    End If
    
    Range("B5") = p_n
    
    ' Se definen los criterios de paro
    If opcion_tol_new.Value = True Then
        Range("B3") = "Margen de Tolerancia"
        Range("A4") = "Tolerancia"
        Range("B4") = Val(tol_new.Value)
        criterio = 1
    Else
        Range("B3") = "Máximo de iteraciones"
        Range("A4") = "Iteraciones"
        Range("B4") = Val(iter_new.Value)
        criterio = 2
    End If
    
    ' 2. Escribir Encabezados e Iteración 0
    Iteración = 0
    Range("A7:C7") = Array("Iteración", "p_n", "|f(p_n)|")
    Cells(8, 1) = Iteración
    Cells(8, 2) = p_n
    Cells(8, 3) = Abs(fun1(p_n))
    
    ' 3. Bucle de cálculo
    Select Case criterio
        Case 1 ' Por Tolerancia
            Do While Abs(fun1(p_n)) > Val(tol_new.Value)
                Iteración = Iteración + 1
                
                ' Aplicar fórmula de Newton-Raphson
                p_n = p_n - (fun1(p_n) / derivada(p_n))
                
                ' Imprimir fila (La fila siempre será Iteración + 8. Ej: Iteración 1 + 8 = Fila 9)
                Cells(Iteración + 8, 1) = Iteración
                Cells(Iteración + 8, 2) = p_n
                Cells(Iteración + 8, 3) = Abs(fun1(p_n))
                
                ' Seguro antidesbordamiento opcional
                If Iteración > 10000 Then Exit Do
            Loop
            
        Case 2 ' Por número de iteraciones
            For i = 1 To Val(iter_new.Value)
                Iteración = Iteración + 1
                
                p_n = p_n - (fun1(p_n) / derivada(p_n))
                
                Cells(Iteración + 8, 1) = Iteración
                Cells(Iteración + 8, 2) = p_n
                Cells(Iteración + 8, 3) = Abs(fun1(p_n))
            Next i
    End Select
    
    ' Se resalta la celda que contiene el resultado de la última iteración
    Range("B" & Iteración + 8).Interior.Color = RGB(117, 219, 255)
    
    ' Se da formato a las tablas generadas
    Range("A7:C" & Iteración + 8).Borders.LineStyle = xlContinuous
    Range("A1:B5").Borders.LineStyle = xlContinuous
    
    ' Se ajusta el tamaño de las celdas
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select ' Para quitar la selección global y dejar la vista limpia
        
End Sub

Private Sub siguiente_newm_Click()
    Dim a As Double
    Dim b As Double
    Dim p_n As Double
    Dim Tolerancia As Double
    Dim Iteracion As Long
    Dim i As Long
    Dim criterio As Integer
    Dim denominador As Double
    
    Sheets("Hoja1").Select
    Cells.Clear
    
    '========================================
    ' LECTURA DE DATOS
    '========================================
    
    a = Val(inferior_new_m.Value)
    b = Val(superior_new_m.Value)
    
    ' Verificar existencia de raíz en [a,b]
'    If fun1(a) * fun1(b) > 0 Then
'        MsgBox "No se halló raíz alguna en el intervalo creado por a = " & a & _
'               " y b = " & b & "." & vbCrLf & _
'               "Ingrese nuevos valores para ambos límites.", vbExclamation
'
'        inferior_new_m.Value = ""
'        superior_new_m.Value = ""
'        tol_new_m.Value = ""
'        iter_new_m.Value = ""
'        aprox_i_new_m.Value = ""
'
'        inferior_new_m.SetFocus
'        Exit Sub
'    End If
    
    '========================================
    ' INFORMACIÓN GENERAL
    '========================================
    
    Range("A1") = "Límite Inferior"
    Range("A2") = "Límite Superior"
    Range("A3") = "Criterio"
    Range("A5") = "Aprox. Inicial"
    
    Range("B1") = a
    Range("B2") = b
    
    ' Aproximación inicial
    If aprox_i_pm_new_m.Value = True Then
        p_n = (a + b) / 2
    Else
        p_n = Val(aprox_i_new_m.Value)
    End If
    
    Range("B5") = p_n
    
    '========================================
    ' CRITERIO DE PARO
    '========================================
    
    If opcion_tol_new_m.Value = True Then
        Range("B3") = "Margen de Tolerancia"
        Range("A4") = "Tolerancia"
        Range("B4") = Val(tol_new_m.Value)
        criterio = 1
        Tolerancia = Val(tol_new_m.Value)
    Else
        Range("B3") = "Máximo de Iteraciones"
        Range("A4") = "Iteraciones"
        Range("B4") = Val(iter_new_m.Value)
        criterio = 2
    End If
    
    '========================================
    ' ENCABEZADOS
    '========================================
    
    Iteracion = 0
    
    Range("A7:D7") = Array("Iteración", "p_n", "f(p_n)", "|f(p_n)|")
    
    Cells(8, 1) = Iteracion
    Cells(8, 2) = p_n
    Cells(8, 3) = fun1(p_n)
    Cells(8, 4) = Abs(fun1(p_n))
    
    '========================================
    ' MÉTODO DE NEWTON MODIFICADO
    '
    ' p(n+1) = p(n) - [f(x)f'(x)] /
    '          ([f'(x)]^2 - f(x)f''(x))
    '========================================
    
    Select Case criterio
        
        '------------------------------------
        ' CRITERIO POR TOLERANCIA
        '------------------------------------
        Case 1
            
            Do While Abs(fun1(p_n)) > Tolerancia
                
                denominador = derivada(p_n) ^ 2 - _
                              fun1(p_n) * segunda_derivada(p_n)
                
                ' Evitar división entre cero
                If denominador = 0 Then
                    MsgBox "El denominador se hizo cero en la iteración " & _
                           Iteracion & ".", vbCritical
                    Exit Do
                End If
                
                Iteracion = Iteracion + 1
                
                p_n = p_n - _
                      (fun1(p_n) * derivada(p_n)) / denominador
                
                Cells(Iteracion + 8, 1) = Iteracion
                Cells(Iteracion + 8, 2) = p_n
                Cells(Iteracion + 8, 3) = fun1(p_n)
                Cells(Iteracion + 8, 4) = Abs(fun1(p_n))
                
                ' Seguro anti-bucle infinito
                If Iteracion > 10000 Then Exit Do
                
            Loop
        
        '------------------------------------
        ' CRITERIO POR NÚMERO DE ITERACIONES
        '------------------------------------
        Case 2
            
            For i = 1 To Val(iter_new_m.Value)
                
                denominador = derivada(p_n) ^ 2 - _
                              fun1(p_n) * segunda_derivada(p_n)
                
                If denominador = 0 Then
                    MsgBox "El denominador se hizo cero en la iteración " & _
                           Iteracion & ".", vbCritical
                    Exit For
                End If
                
                Iteracion = Iteracion + 1
                
                p_n = p_n - _
                      (fun1(p_n) * derivada(p_n)) / denominador
                
                Cells(Iteracion + 8, 1) = Iteracion
                Cells(Iteracion + 8, 2) = p_n
                Cells(Iteracion + 8, 3) = fun1(p_n)
                Cells(Iteracion + 8, 4) = Abs(fun1(p_n))
                
            Next i
    End Select
    
    '========================================
    ' FORMATO FINAL
    '========================================
    
    ' Resaltar resultado final
    Range("B" & Iteracion + 8).Interior.Color = RGB(117, 219, 255)
    
    ' Bordes
    Range("A7:D" & Iteracion + 8).Borders.LineStyle = xlContinuous
    Range("A1:B5").Borders.LineStyle = xlContinuous
    
    ' Ajustar columnas
    Cells.EntireColumn.AutoFit
    
    ' Quitar selección
    Range("A1").Select

End Sub

Private Sub siguiente_sec_Click()
    Dim a As Double
    Dim b As Double
    Dim Tolerancia As Double
    Dim n_iteración As Double
    Dim c As Double
    Dim Iteración As Long
    Dim i As Double
    Dim criterio As Integer
        
    Sheets("Hoja1").Select
    Cells.Clear
    
    If fun1(inferior_sec.Value) * fun1(superior_sec.Value) > 0 Then 'Se verifica la existencia de una raíz de la función en el intervalo creado por a y b
        MsgBox "No se halló raíz alguna en el intervalo creado por a = " & a & " y b = " & b & ", ingrese nuevos valores para ambos límites"
        inferior_sec.Value = ""
        superior_sec.Value = ""
        tol_sec.Value = ""
        iter_sec.Value = ""
        inferior_sec.SetFocus
        Exit Sub
    End If
    
    Range("A1") = "Límite Inferior"
    Range("A2") = "Límite Superior"
    Range("A3") = "Criterio"
    Range("B1") = inferior_sec.Value
    Range("B2") = superior_sec.Value
    
    a = inferior_sec.Value
    b = superior_sec.Value
    c = b - ((fun1(b) * (b - a)) / (fun1(b) - fun1(a)))
    Iteración = 0
    
    Range("A6:C6") = Array("Iteración", "p_n", "|f(p_n)|")
    
    If opcion_tol_sec.Value = True Then
        Range("B3") = "Margen de Tolerancia"
        Range("A4") = "Tolerancia"
        Range("B4") = tol_sec.Value
        criterio = 1
    Else
        Range("B3") = "Máximo de iteraciones"
        Range("A4") = "Iteraciones"
        Range("B4") = iter_sec.Value
        criterio = 2
    End If
    
    Select Case criterio
        Case 1
        Do While Abs(fun1(c)) > Val(tol_sec.Value)
            c = b - ((fun1(b) * (b - a)) / (fun1(b) - fun1(a)))
            Iteración = Iteración + 1
            Cells(Iteración + 6, 1) = Iteración
            Cells(Iteración + 6, 2) = c
            Cells(Iteración + 6, 3) = Abs(fun1(c))
            a = b
            b = c
        Loop
        
        Case 2
        For i = 1 To Val(iter_sec.Value)
            c = b - ((fun1(b) * (b - a)) / (fun1(b) - fun1(a)))
            Iteración = Iteración + 1
            Cells(Iteración + 6, 1) = Iteración
            Cells(Iteración + 6, 2) = c
            Cells(Iteración + 6, 3) = Abs(fun1(c))
            a = b
            b = c
        Next i
    End Select
        
    'Se resalta la celda que contiene el resultado de la última iteración
    Range("B" & Iteración + 6).Interior.Color = RGB(117, 219, 255)
    
    'Se da formato de tabla a las tablas generadas
    Range("A6:C" & Iteración + 6).Borders.LineStyle = xlContinuous
    Range("A1:B4").Borders.LineStyle = xlContinuous
    
    'Se ajusta el tamaño de las celdas al contenido de estas
    Cells.Select
    Cells.EntireColumn.AutoFit
        
End Sub

Private Sub siguiente2_Click()
    If falsi.Value = True Then
        UserForm1.MultiPage1.Value = 1
    ElseIf newton.Value = True Then
        UserForm1.MultiPage1.Value = 2
    ElseIf bisección.Value = True Then
        UserForm1.MultiPage1.Value = 3
    ElseIf secante.Value = True Then
        UserForm1.MultiPage1.Value = 4
    ElseIf newton_m.Value = True Then
        UserForm1.MultiPage1.Value = 5
    End If
End Sub

Private Sub superior_bis_Change()
    validar_bis
End Sub

Private Sub superior_Change()
    validar
End Sub

Private Sub superior_new_Change()
    validar_new
End Sub

Private Sub superior_new_m_Change()
    validar_new_m
End Sub

Private Sub superior_sec_Change()
    validar_sec
End Sub

Private Sub TextBox4_Change()

End Sub

Private Sub TextBox5_Change()

End Sub

Private Sub tol_bis_Change()
    validar_bis
End Sub

Private Sub tol_Change()
    validar
End Sub

Private Sub tol_new_Change()
    validar_new
End Sub

Private Sub tol_new_m_Change()
    validar_new_m
End Sub

Private Sub tol_sec_Change()
    validar_sec
End Sub

Private Sub UserForm_Initialize()

    'Hace que la macro inicie en la pestaña de selección de método
    MultiPage1.Value = 0
    'Deshabilitar el botón siguiente
    siguiente.Enabled = False
    siguiente2.Enabled = False
    siguiente_bis.Enabled = False
    siguiente_sec.Enabled = False
    siguiente_new.Enabled = False
    siguiente_newm.Enabled = False
    
    'Ocultar etiquetas y cuadros de texto
    label_tol.Visible = False
    label_tol_bis.Visible = False
    label_tol_sec.Visible = False
    label_tol_new.Visible = False
    label_tol_new_m.Visible = False
    label_iter.Visible = False
    label_iter_bis.Visible = False
    label_iter_sec.Visible = False
    label_iter_new.Visible = False
    label_iter_new_m.Visible = False
    tol.Visible = False
    tol_bis.Visible = False
    tol_sec.Visible = False
    tol_new.Visible = False
    tol_new_m.Visible = False
    iter.Visible = False
    iter_bis.Visible = False
    iter_sec.Visible = False
    iter_new_m.Visible = False
    iter_new.Visible = False
    aprox_i_new_m.Visible = False
    aprox_i.Visible = False
    
    'Para que no aparezcan seleccionados los botones de acción
    opcion_tolerancia.Value = 0
    opcion_iteraciones.Value = 0
    opcion_tol_bis.Value = 0
    opcion_tol_sec.Value = 0
    opcion_tol_new.Value = 0
    opcion_tol_new_m.Value = 0
    opcion_iter_bis.Value = 0
    opcion_iter_sec.Value = 0
    opcion_iter_new.Value = 0
    opcion_iter_new_m.Value = 0
    
    'Para que no se seleccione ningún método al iniciar la macro
    newton.Value = False
    falsi.Value = False
    bisección.Value = False
    secante.Value = False
    newton_m.Value = False
    
End Sub

Private Sub validar()
    Dim esValido As Boolean
    esValido = True ' Asumimos que todo está bien hasta que se demuestre lo contrario
    
    ' 1. Validar que los límites inferior y superior sean estrictamente números
    If inferior.Value = "" Or Not IsNumeric(inferior.Value) Then esValido = False
    If superior.Value = "" Or Not IsNumeric(superior.Value) Then esValido = False
    
    ' 2. Validar las condiciones especiales de paro
    If opcion_tolerancia.Value = True Then
        
        ' Caso A: Se eligió Tolerancia
        If tol.Value = "" Or Not IsNumeric(tol.Value) Then
            esValido = False
        ElseIf Val(tol.Value) <= 0 Or Val(tol.Value) >= 1 Then
            ' La tolerancia debe ser estrictamente mayor a 0 y menor a 1 (ej. 0.001)
            esValido = False
        End If
        
    ElseIf opcion_iteraciones.Value = True Then
        
        ' Caso B: Se eligieron Iteraciones
        If iter.Value = "" Or Not IsNumeric(iter.Value) Then
            esValido = False
        ElseIf Val(iter.Value) < 1 Or Val(iter.Value) > 100000 Then
            ' Límite anti-desbordamiento: máximo 100,000 iteraciones.
            ' Si un método no converge en 100,000 ciclos, es mejor detenerlo para no congelar Excel.
            esValido = False
        ElseIf InStr(1, iter.Value, ".") > 0 Or InStr(1, iter.Value, ",") > 0 Then
            ' Evitar que metan decimales en las iteraciones (no existe la iteración 1.5)
            esValido = False
        End If
        
    Else
        ' Si el usuario no ha seleccionado ningún botón de opción (Tolerancia o Iteraciones)
        esValido = False
    End If
    
    ' Finalmente, el botón Siguiente solo se enciende si TODO fue correcto
    siguiente.Enabled = esValido
End Sub

Private Sub validar_bis()
    Dim esValido As Boolean
    esValido = True ' Asumimos que todo está bien hasta que se demuestre lo contrario
    
    ' 1. Validar que los límites inferior y superior sean estrictamente números
    If inferior_bis.Value = "" Or Not IsNumeric(inferior_bis.Value) Then esValido = False
    If superior_bis.Value = "" Or Not IsNumeric(superior_bis.Value) Then esValido = False
    
    ' 2. Validar las condiciones especiales de paro
    If opcion_tol_bis.Value = True Then
        
        ' Caso A: Se eligió Tolerancia
        If tol_bis.Value = "" Or Not IsNumeric(tol_bis.Value) Then
            esValido = False
        ElseIf Val(tol_bis.Value) <= 0 Or Val(tol_bis.Value) >= 1 Then
            ' La tolerancia debe ser estrictamente mayor a 0 y menor a 1 (ej. 0.001)
            esValido = False
        End If
        
    ElseIf opcion_iter_bis.Value = True Then
        
        ' Caso B: Se eligieron Iteraciones
        If iter_bis.Value = "" Or Not IsNumeric(iter_bis.Value) Then
            esValido = False
        ElseIf Val(iter_bis.Value) < 1 Or Val(iter_bis.Value) > 100000 Then
            ' Límite anti-desbordamiento: máximo 100,000 iteraciones.
            ' Si un método no converge en 100,000 ciclos, es mejor detenerlo para no congelar Excel.
            esValido = False
        ElseIf InStr(1, iter_bis.Value, ".") > 0 Or InStr(1, iter_bis.Value, ",") > 0 Then
            ' Evitar que metan decimales en las iteraciones (no existe la iteración 1.5)
            esValido = False
        End If
        
    Else
        ' Si el usuario no ha seleccionado ningún botón de opción (Tolerancia o Iteraciones)
        esValido = False
    End If
    
    ' Finalmente, el botón Siguiente solo se enciende si TODO fue correcto
    siguiente_bis.Enabled = esValido
End Sub

Private Sub validar_sec()
    Dim esValido As Boolean
    esValido = True ' Asumimos que todo está bien hasta que se demuestre lo contrario
    
    ' 1. Validar que los límites inferior y superior sean estrictamente números
    If inferior_sec.Value = "" Or Not IsNumeric(inferior_sec.Value) Then esValido = False
    If superior_sec.Value = "" Or Not IsNumeric(superior_sec.Value) Then esValido = False
    
    ' 2. Validar las condiciones especiales de paro
    If opcion_tol_sec.Value = True Then
        
        ' Caso A: Se eligió Tolerancia
        If tol_sec.Value = "" Or Not IsNumeric(tol_sec.Value) Then
            esValido = False
        ElseIf Val(tol_sec.Value) <= 0 Or Val(tol_sec.Value) >= 1 Then
            ' La tolerancia debe ser estrictamente mayor a 0 y menor a 1 (ej. 0.001)
            esValido = False
        End If
        
    ElseIf opcion_iter_sec.Value = True Then
        
        ' Caso B: Se eligieron Iteraciones
        If iter_sec.Value = "" Or Not IsNumeric(iter_sec.Value) Then
            esValido = False
        ElseIf Val(iter_sec.Value) < 1 Or Val(iter_sec.Value) > 100000 Then
            ' Límite anti-desbordamiento: máximo 100,000 iteraciones.
            ' Si un método no converge en 100,000 ciclos, es mejor detenerlo para no congelar Excel.
            esValido = False
        ElseIf InStr(1, iter_sec.Value, ".") > 0 Or InStr(1, iter_sec.Value, ",") > 0 Then
            ' Evitar que metan decimales en las iteraciones (no existe la iteración 1.5)
            esValido = False
        End If
        
    Else
        ' Si el usuario no ha seleccionado ningún botón de opción (Tolerancia o Iteraciones)
        esValido = False
    End If
    
    ' Finalmente, el botón Siguiente solo se enciende si TODO fue correcto
    siguiente_sec.Enabled = esValido
End Sub

Private Sub validar_new()
    Dim esValido As Boolean
    esValido = True ' Asumimos que todo está bien hasta que se demuestre lo contrario
    
    ' 1. Validar que los límites inferior y superior sean estrictamente números
    If inferior_new.Value = "" Or Not IsNumeric(inferior_new.Value) Then esValido = False
    If superior_new.Value = "" Or Not IsNumeric(superior_new.Value) Then esValido = False
    
    ' 2. Validar las condiciones especiales de paro
    If opcion_tol_new.Value = True Then
        If tol_new.Value = "" Or Not IsNumeric(tol_new.Value) Then
            esValido = False
        ElseIf Val(tol_new.Value) <= 0 Or Val(tol_new.Value) >= 1 Then
            esValido = False
        End If
    ElseIf opcion_iter_new.Value = True Then
        If iter_new.Value = "" Or Not IsNumeric(iter_new.Value) Then
            esValido = False
        ElseIf Val(iter_new.Value) < 1 Or Val(iter_new.Value) > 100000 Then
            esValido = False
        ElseIf InStr(1, iter_new.Value, ".") > 0 Or InStr(1, iter_new.Value, ",") > 0 Then
            esValido = False
        End If
    Else
        ' Si no elige ni tolerancia ni iteraciones
        esValido = False
    End If
    
    ' 3. Validar la aproximación inicial
    If aprox_i_m.Value = True Then
        ' Si eligió manual, la caja de texto debe ser un número dentro del rango
        If aprox_i.Value = "" Or Not IsNumeric(aprox_i.Value) Then
            esValido = False
        ElseIf Val(aprox_i.Value) < Val(inferior_new.Value) Or Val(aprox_i.Value) > Val(superior_new.Value) Then
            esValido = False
        End If
    ElseIf aprox_i_pm.Value = True Then
        ' Si eligió punto medio, es válido, no revisamos cajas extra
    Else
        ' Si no ha seleccionado ninguna forma de aproximación inicial
        esValido = False
    End If
        
    ' Finalmente, el botón Siguiente solo se enciende si TODO fue correcto
    siguiente_new.Enabled = esValido
End Sub
Private Sub validar_new_m()
    Dim esValido As Boolean
    esValido = True ' Asumimos que todo está bien hasta que se demuestre lo contrario
    
    ' 1. Validar que los límites inferior y superior sean estrictamente números
    If inferior_new_m.Value = "" Or Not IsNumeric(inferior_new_m.Value) Then esValido = False
    If superior_new_m.Value = "" Or Not IsNumeric(superior_new_m.Value) Then esValido = False
    
    ' 2. Validar las condiciones especiales de paro
    If opcion_tol_new_m.Value = True Then
        If tol_new_m.Value = "" Or Not IsNumeric(tol_new_m.Value) Then
            esValido = False
        ElseIf Val(tol_new_m.Value) <= 0 Or Val(tol_new_m.Value) >= 1 Then
            esValido = False
        End If
    ElseIf opcion_iter_new_m.Value = True Then
        If iter_new_m.Value = "" Or Not IsNumeric(iter_new_m.Value) Then
            esValido = False
        ElseIf Val(iter_new_m.Value) < 1 Or Val(iter_new_m.Value) > 100000 Then
            esValido = False
        ElseIf InStr(1, iter_new_m.Value, ".") > 0 Or InStr(1, iter_new_m.Value, ",") > 0 Then
            esValido = False
        End If
    Else
        ' Si no elige ni tolerancia ni iteraciones
        esValido = False
    End If
    
    ' 3. Validar la aproximación inicial
    If aprox_i_m_new_m.Value = True Then
        ' Si eligió manual, la caja de texto debe ser un número dentro del rango
        If aprox_i_new_m.Value = "" Or Not IsNumeric(aprox_i_new_m.Value) Then
            esValido = False
        ElseIf Val(aprox_i_new_m.Value) < Val(inferior_new_m.Value) Or Val(aprox_i_new_m.Value) > Val(superior_new_m.Value) Then
            esValido = False
        End If
    ElseIf aprox_i_pm_new_m.Value = True Then
        ' Si eligió punto medio, es válido, no revisamos cajas extra
    Else
        ' Si no ha seleccionado ninguna forma de aproximación inicial
        esValido = False
    End If
        
    ' Finalmente, el botón Siguiente solo se enciende si TODO fue correcto
    siguiente_newm.Enabled = esValido
End Sub

