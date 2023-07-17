' En esta funcion podemos utilizar para cifrar una frase a

Function CifrarFrase(frase As String, clave As String) As String
    Dim resultado As String
    Dim i As Integer
    Dim j As Integer
    Dim arregloSeparado() As String
    Dim totalFrase As Integer
    Dim totalClave As Integer
    totalFrase = Len(frase)
    totalClave = Len(clave)
    Dim numeroArreglos As Double
    numeroArreglos = totalFrase / totalClave
    Dim totalArreglos As Integer
    totalArreglos = Application.WorksheetFunction.Ceiling(numeroArreglos, 1) ' Redondea siempre hacia arriba al múltiplo de 1
    ReDim arregloSeparado(totalArreglos, 1)
    Dim fraseFormada As String
    Dim letra As String
    j = 1
    
    For i = 1 To totalFrase
        letra = Chr(Asc(Mid(frase, i, 1)))
        If ((i Mod totalClave) = 0) Then
        fraseFormada = fraseFormada & letra
        arregloSeparado(j, 1) = fraseFormada
        ' resultado = resultado & arregloSeparado(j, 1)
        j = j + 1
        fraseFormada = ""
        Else
        fraseFormada = fraseFormada & letra
        End If
        If j >= totalArreglos And i = totalFrase Then
            arregloSeparado(j, 1) = fraseFormada
            ' resultado = resultado & arregloSeparado(j, 1)
        End If
    Next i
    
    Dim x As Integer
    Dim y As Integer
    Dim arregloCifrado() As String
    ReDim arregloCifrado(totalClave, 1)
    Dim claveFormada As String
    Dim ordenClave(6) As Integer
    ordenClave(0) = 4
    ordenClave(1) = 5
    ordenClave(2) = 3
    ordenClave(3) = 6
    ordenClave(4) = 2
    ordenClave(5) = 7
    ordenClave(6) = 1
    

    Dim numero As Integer
    
    For x = 1 To totalClave
        claveFormada = "" ' Reiniciar la variable claveFormada para cada iteración de x
        fraseFormada = ""
        For y = 1 To totalArreglos
            fraseFormada = arregloSeparado(y, 1)
            letra = Mid(fraseFormada, x, 1)
            claveFormada = claveFormada & letra
        Next y
        arregloCifrado(x, 1) = claveFormada ' Asignar claveFormada a arregloCifrado(x, 1)
        ' resultado = resultado & arregloCifrado(x, 1)
    Next x
    
    For x = 0 To totalClave - 1
        numero = ordenClave(x)
        resultado = resultado & arregloCifrado(numero, 1)
    Next x
   
    
    CifrarFrase = resultado
End Function
