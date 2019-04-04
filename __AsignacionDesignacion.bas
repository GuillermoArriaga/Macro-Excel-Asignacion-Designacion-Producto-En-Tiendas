Attribute VB_Name = "Módulo1"
' ==================================================================================='
'                                                                                    '
' Creado por Guillermo Arriaga García  del 17/07/2016 al 14/08/2016 en 55 hrs aprox. '
'                                                                                    '
' Libro de excel programado para calcular la asignación de producto disponible y la  '
'    designación de pedido para que las tiendas puedan seguir vendiendo lo que han   '
'    vendido.                                                                        '
'                                                                                    '
' Contacto: guillermoarriagag@gmail.com                                              '
'                                                                                    '
' ==================================================================================='

Sub Informacion()
    
    Dim hi As VbMsgBoxResult
    hi = MsgBox("Este libro indica como asignar el producto que hay por talla en cada tienda ofreciendo tres procedimientos: " & vbCrLf & "+ ASIGNAR PRODUCTO: Llena rápido el producto disponible buscando que todas las tasas no sean menores a la min deseada. Puede sobrar producto." & vbCrLf & "+ TODO: Asigna todo el producto sin importar algun limite de tasa." & vbCrLf & "+ 1X1: Asigna según tasa min el producto de uno por uno, de modo que es lento. Es útil si se quiere comparar resultados con el rápido.", vbOKOnly + vbInformation)
    hi = MsgBox("IMPORTANTE" & vbCrLf & "Se deben llenar los datos numéricos con números: Tiene, Vendió, Disponible, Tasa min, Tallas, Tiendas, Disponible, Tasa min, Tasa med y Tasa max." & vbCrLf & vbCrLf & "Siempre deben estar los datos de cantidad de tallas, tiendas, tasa min (azul) y las tasas para el pedido: min, med y max.", vbOKOnly + vbInformation)
    hi = MsgBox("El programa de este libro trabaja con cualquier cantidad de tiendas y tallas, mientras la multiplicación  de estas sea menor a 1'048,576." & vbCrLf & "No es necesario agregar alguna formula, el programa ya las incluye y no se pueden modificar mas que en el código fuente." & vbCrLf & "El botón de LLENAR PRODUCTO asigna por talla el producto disponible, A26, y limpia lo asignado." & vbCrLf & "El botón CREAR DATOS hace un llenado aleatorio de mercancía, venta y disponible.", vbOKOnly + vbInformation)
    hi = MsgBox("FORMULAS UTILIZADAS: " & vbCrLf & "IMPACTO = SI(Vendió > 0, 100/Vendió,100/0.667) Indica el aumento en tasa al asignar un producto. " & vbCrLf & "TASA INICIAL = Tiene*Impacto" & vbCrLf & "TASA = Asignado * Impacto + TasaInicial." & vbCrLf & "El objetivo que se sigue es asignar producto aumentando las tasas pequeñas de modo que tiendan a igualarse en lo posible. Una tasa de 100 equivale a poder vender lo que se ha vendido antes." & vbCrLf & "PEDIDO = ENTERO( (Tasa deseada - TasaActual)/Impacto + 0.5 )", vbOKOnly + vbInformation)
    hi = MsgBox("Este libro también calcula el pedido para alcanzar ciertas tasas deseadas 100, 150, 200 por default pero ajustables según se desee. Las fórmulas de tasa y pedido quedan activas por si se quiere hacer cambios de asignación, sólo que estos cambios no se registrarán automáticamente en la casilla de disponible por talla, tendría que ser manual." & vbCrLf & "El pedido por talla YA TIENE RESTADO el disponible sobrante, es decir, si para la tasa deseada 150 se requieren 300 productos y sobraron de la asignación 10, entonces indica 290... así en el pedido por talla min, med y max." & vbCrLf & " En la segunda hoja se muestra un resumen de disponible y de los tres pedidos por cada talla.", vbOKOnly + vbInformation)

End Sub

Sub CrearDatos()
   
   Dim numTallas As Long
   Dim numTiendas As Integer
   Dim totalFilas As Long
   Dim talla As Integer
   Dim tienda As Integer
   Dim h As Long
   Dim v As Integer
   Dim numProducto As Integer
   Dim conta As Integer
   
'Tiempo en cero
   Cells(16, 1) = 0
   
'Recepcion de datos
   If Cells(22, 1) = "" Then
      numTallas = CInt(InputBox("Se comenzará un llenado aleatorio de datos." & vbCrLf & "Ingrese la cantidad de tallas."))
      Cells(22, 1) = numTallas
   Else
      numTallas = Cells(22, 1)
   End If
   
   If Cells(24, 1) = "" Then
      numTiendas = CInt(InputBox("Ingrese la cantidad de tiendas."))
      Cells(24, 1) = numTiendas
   Else
      numTiendas = Cells(24, 1)
   End If
   
'Paso por cada fila
   totalFilas = numTallas * numTiendas + 1
   talla = 1
   conta = 1
   Cells(2, 6) = Int(200 * Rnd())

   For h = 2 To totalFilas
      tienda = conta
      If tienda = numTiendas + 1 Then
         talla = talla + 1
         Cells(h, 6) = Int(200 * Rnd())
         conta = 1
         tienda = conta
      End If
      Cells(h, 2) = tienda
      Cells(h, 3) = talla
      Cells(h, 4) = Int(200 * Rnd())
      Cells(h, 5) = Int(200 * Rnd())
      conta = conta + 1
   Next h
   
Exit Sub

cancelacion:
   MsgBox "Cancelacion de macro"
End Sub

Sub LlenarProducto()
   
   Dim numTallas As Long
   Dim producto As Integer
   Dim numTiendas As Integer
   Dim totalFilas As Long
   Dim fila As Long

Sheets("Asignacion").Select
    
Application.Calculation = xlManual
    
    producto = Cells(26, 1)
    numTiendas = Cells(24, 1)
    numTallas = Cells(22, 1)
    totalFilas = (numTiendas * numTallas + 1)

'Borra producto y asignacion
Sheets("Asignacion").Select
    Range("F2:G" & totalFilas).Select
    Selection.ClearContents
    Range("F2").Select

'Asigna valor a cada primero de talla
   numTallas = numTallas - 1
   For fila = 0 To numTallas
      Cells(fila * numTiendas + 2, 6) = producto
   Next fila
   numTallas = numTallas + 1
   
'Regresa al principio
    Range("F2").Select

Application.Calculation = xlAutomatic

End Sub

Sub AsignacionComplejaHastaTasa()
   
   Dim numTallas As Long
   Dim hMin() As Long
   Dim cTalla As Long
   Dim h1 As Long
   Dim h As Long
   
   Dim v As Integer
   Dim pos As Integer
   Dim colTasa As Integer
   Dim colTasaIni As Integer
   Dim colImpacto As Integer
   Dim colDisponible As Integer
   Dim colAsignado As Integer
   Dim numTiendas As Integer
   Dim revision As Integer
   Dim tasaMin As Integer
   Dim cTienda As Integer
   Dim cMin As Integer
   Dim pTas As Integer
   
   Dim avance As Double
   Dim EndTime As Double
   Dim producto As Double
   Dim StartTime As Double
   Dim sumaAsigna As Double
   Dim asigna() As Double
   Dim tasas() As Double
   Dim tasaM As Double
   Dim swap As Double
   
   Dim tasaAlcanzada As Boolean
   
'Asignacion directa una tienda
   If Cells(24, 1) = 1 Then
      Call CasoUnaTienda
      Exit Sub
   End If
   
StartTime = Timer
Sheets("Asignacion").Select
   Call DatosIniciales

Application.Calculation = xlManual

' Capturar valores
   Cells(15, 1) = 0
   Cells(16, 1) = 0
   h = 2
   v = 8   ' Inicia en tasa del primero del grupo
   colTasa = v
   colAsignado = v - 1
   colTasaIni = v + 2
   colImpacto = v + 1
   colDisponible = v - 2
   
   tasaM = Cells(20, 1)
   numTiendas = Cells(24, 1)
   numTallas = Cells(22, 1)
   revision = numTiendas - 1
   avance = 50 / numTallas
   
   ReDim hMin(numTiendas)
   ReDim asigna(numTiendas)
   ReDim tasas(numTiendas)
   
   conta = 0  ' para imprimir avance de 50 en 50 tallas
   
'= PASO POR CADA TALLA
   For cTalla = 1 To numTallas
   
   'Si nada hay para asignar pasa a la siguiente talla
      Cells(h, colDisponible) = Int(Cells(h, colDisponible) + 0.5)
      If Cells(h, colDisponible) = 0 Then GoTo siguienteTalla
      
   'Orden de tasas
      For cMin = 1 To numTiendas
         hMin(cMin) = h + cMin - 1
         tasas(cMin) = Cells(h + cMin - 1, colTasa)
         asigna(cMin) = 0
      Next cMin
      For min = 1 To numTiendas
         pos = min - 1
         For cMin = (min + 1) To numTiendas
            If tasas(cMin) < tasas(pos + 1) Then
               pos = cMin - 1
            End If
         Next cMin
         swap = tasas(min)          ' Intercambio usando una variable double
         tasas(min) = tasas(pos + 1)
         tasas(pos + 1) = swap
         h1 = hMin(min)             ' Intercambio usando una variable long
         hMin(min) = hMin(pos + 1)
         hMin(pos + 1) = h1
      Next min
   
   'Sgt talla si ninguna tasa es menor a la pedida
      'If Cells(hMin(1), colTasa) >= tasaM Then
      If tasas(1) >= tasaM Then
         'tasaAlcanzada = True
         GoTo siguienteTalla
      Else
         tasaAlcanzada = False
      End If
      
   'Ubicacion de posiciones menores a la tasa min deseada
      pTas = numTiendas
      For cMin = 1 To numTiendas
         If tasas(cMin) > tasaM Then
            'pTas = Int(hMin(cMin - 1) - h)
            pTas = cMin - 1
            'Cells(hMin(cMin), colTasa) = tasaM
            Exit For
         End If
      Next cMin
   ' hMin(pTas) es la ultima posicion menor
      
   'Deteccion de producto para asignar
      producto = Cells(h, colDisponible)
      
'= REVISIÓN POR TASAS DE LAS TIENDAS MENORES A LAS MAYORES
      For cTienda = 1 To numTiendas
   
'==== IDENTIFICACION Y ASIGNACION DE PRODUCTO
         sumaAsigna = 0
      'Identificacion del producto necesario para llegar a siguiente tasa
         If cTienda = numTiendas Then
         'Se fija cuanto asignaria para alcanzar a tasaM si ya estamos asognando a todas las tiendas
            For cMin = 1 To cTienda
               asigna(cMin) = (tasaM - Cells(hMin(cMin), colTasa)) / Cells(hMin(cMin), colImpacto)
               sumaAsigna = sumaAsigna + asigna(cMin)
            Next cMin
         Else
         'Tasas menores a la de la posicion cTienda + 1 le alcanzan
            If tasaM < Cells(hMin(cTienda + 1), colTasa) Then
               For cMin = 1 To cTienda
                  asigna(cMin) = (tasaM - Cells(hMin(cMin), colTasa)) / Cells(hMin(cMin), colImpacto)
                  sumaAsigna = sumaAsigna + asigna(cMin)
               Next cMin
               tasaAlcanzada = True
            Else
               For cMin = 1 To cTienda
                  asigna(cMin) = (Cells(hMin(cTienda + 1), colTasa) - Cells(hMin(cMin), colTasa)) / Cells(hMin(cMin), colImpacto)
                  sumaAsigna = sumaAsigna + asigna(cMin)
               Next cMin
               tasaAlcanzada = False
            End If
         End If
      
      'Asignacion de producto, o todo o la parte correspondiente.
         'Si es posible la asignacion la realiza
         If producto >= sumaAsigna Then
            producto = producto - sumaAsigna
            For cMin = 1 To cTienda
               Cells(hMin(cMin), colAsignado) = asigna(cMin) + Cells(hMin(cMin), colAsignado)
               Cells(hMin(cMin), colTasa) = Cells(hMin(cMin), colTasaIni) + Cells(hMin(cMin), colImpacto) * Cells(hMin(cMin), colAsignado)
            Next cMin
            
         'Elabora resultados si la tasa deseada se ha superado
         'Si no, Continua con el siguiente nivel de asignacion, que se alcance la tasa de la siguiente talla
            If ((producto > 0 And pTas > cTienda) And tasaAlcanzada = False) Then GoTo siguienteNivel
         
         'Si no, entonces reparte lo disponible según lo que le iba a tocar con respecto a lo que se ubiera asignado, así se mantienen tasas semejantes
         Else
            swap = producto
            For cMin = 1 To cTienda
               Cells(hMin(cMin), colAsignado) = producto * asigna(cMin) / sumaAsigna + Cells(hMin(cMin), colAsignado)
               Cells(hMin(cMin), colTasa) = Cells(hMin(cMin), colTasaIni) + Cells(hMin(cMin), colImpacto) * Cells(hMin(cMin), colAsignado)
               swap = swap - producto * asigna(cMin) / sumaAsigna
            Next cMin
            producto = swap
         End If
      
'==== ELABORACION DE RESULTADOS DE ESTA TALLA
      'Redondeo a entero. Primera vez que se resta a lo disponible inicial
         For cMin = 1 To cTienda
            Cells(hMin(cMin), colAsignado) = Int(0.5 + Cells(hMin(cMin), colAsignado))
            Cells(h, colDisponible) = Int(Cells(h, colDisponible) - Cells(hMin(cMin), colAsignado))
            Cells(hMin(cMin), colTasa) = Cells(hMin(cMin), colTasaIni) + Cells(hMin(cMin), colImpacto) * Cells(hMin(cMin), colAsignado)
         Next cMin
         
         If Cells(h, colDisponible) = 0 Then GoTo siguienteTalla 'nada sobro o falto
      
      'Arreglo sobrante por redondeo
         'Busca tasa min
         pos = 1
         For cMin = 2 To numTiendas
            If Cells(hMin(cMin), colTasa) < Cells(hMin(pos), colTasa) Then
               pos = cMin
            End If
         Next cMin
         
         While (Cells(h, colDisponible) > 0 And tasaM > Cells(hMin(pos), colTasa))
         'Asigna un producto
            Cells(hMin(pos), colAsignado) = Cells(hMin(pos), colAsignado) + 1
            Cells(hMin(pos), colTasa) = Cells(hMin(pos), colTasa) + Cells(hMin(pos), colImpacto)
            Cells(h, colDisponible) = Cells(h, colDisponible) - 1
         'Busca sgt tasa min
            pos = 1
            For cMin = 2 To cTienda
               If Cells(hMin(cMin), colTasa) < Cells(hMin(pos), colTasa) Then
                  pos = cMin
               End If
            Next cMin
         Wend
         
      'Arreglo faltante por redondeo, quita hasta dejar producto = 0
         While (Cells(h, colDisponible) < 0)
         'Desasigna un producto
            pos = 1
            For cMin = 2 To cTienda
               If (Cells(hMin(cMin), colTasa) > Cells(hMin(pos), colTasa) And Cells(hMin(pos), colAsignado) > 0) Then
                  pos = cMin
               End If
            Next cMin
            Cells(hMin(pos), colAsignado) = Cells(hMin(pos), colAsignado) - 1
            Cells(hMin(pos), colTasa) = Cells(hMin(pos), colTasa) - Cells(hMin(pos), colImpacto)
            Cells(h, colDisponible) = Cells(h, colDisponible) + 1
         Wend

siguienteTalla:
         Exit For

siguienteNivel:
      Next cTienda
      
   'Fila de la siguiente talla
      h = h + numTiendas
      
   'Impresion de avance
      conta = conta + 1
      If conta = 50 Then
         Cells(15, 1) = Cells(15, 1) + avance
         EndTime = Timer
         Cells(16, 1) = Cells(16, 1) + (EndTime - StartTime)
         StartTime = Timer
         Cells(14, 1).Calculate
         Cells(1, 1).Calculate
         Cells(14, 1).Activate
         conta = 0
      End If
      
'Cambio de talla
   Next cTalla

Application.Calculation = xlAutomatic
   Call ActivarFormulaTasa

'Tiempo de ejecucion
   EndTime = Timer
   Cells(16, 1) = Cells(16, 1) + (EndTime - StartTime)
   Cells(14, 1).Activate
   Cells(15, 1) = 1

End Sub

Sub AsignacionComplejaTodoElProducto()
   
   Dim numTallas As Long
   Dim hMin() As Long
   Dim cTalla As Long
   Dim h1 As Long
   Dim h As Long
   
   Dim v As Integer
   Dim pos As Integer
   Dim colTasa As Integer
   Dim colTasaIni As Integer
   Dim colImpacto As Integer
   Dim colDisponible As Integer
   Dim colAsignado As Integer
   Dim numTiendas As Integer
   Dim revision As Integer
   Dim tasaMin As Integer
   Dim cTienda As Integer
   Dim cMin As Integer
   Dim pTas As Integer
   
   Dim avance As Double
   Dim EndTime As Double
   Dim producto As Double
   Dim StartTime As Double
   Dim sumaAsigna As Double
   Dim asigna() As Double
   Dim tasas() As Double
   Dim swap As Double
   

'Asignacion directa una tienda
   If Cells(24, 1) = 1 Then
      Call CasoUnaTiendaTodo
      Exit Sub
   End If


StartTime = Timer
Sheets("Asignacion").Select
   Call DatosIniciales

Application.Calculation = xlManual

' Capturar valores
   Cells(15, 1) = 0
   Cells(16, 1) = 0
   h = 2
   v = 8   ' Inicia en tasa del primero del grupo
   colTasa = v
   colAsignado = v - 1
   colTasaIni = v + 2
   colImpacto = v + 1
   colDisponible = v - 2
   
   numTiendas = Cells(24, 1)
   numTallas = Cells(22, 1)
   revision = numTiendas - 1
   avance = 50 / numTallas
   
   ReDim hMin(numTiendas)
   ReDim asigna(numTiendas)
   ReDim tasas(numTiendas)
   
   conta = 0  ' para imprimir avance de 50 en 50 tallas
   
'= PASO POR CADA TALLA
   For cTalla = 1 To numTallas
   
   'Si nada hay para asignar pasa a la siguiente talla
      Cells(h, colDisponible) = Int(Cells(h, colDisponible) + 0.5)
      If Cells(h, colDisponible) = 0 Then GoTo siguienteTalla
      
   'Orden de tasas
      For cMin = 1 To numTiendas
         hMin(cMin) = h + cMin - 1
         tasas(cMin) = Cells(h + cMin - 1, colTasa)
         asigna(cMin) = 0
      Next cMin
      For min = 1 To numTiendas
         pos = min - 1
         For cMin = (min + 1) To numTiendas
            If tasas(cMin) < tasas(pos + 1) Then
               pos = cMin - 1
            End If
         Next cMin
         swap = tasas(min)          ' Intercambio usando una variable double
         tasas(min) = tasas(pos + 1)
         tasas(pos + 1) = swap
         h1 = hMin(min)             ' Intercambio usando una variable long
         hMin(min) = hMin(pos + 1)
         hMin(pos + 1) = h1
      Next min
   
   'Deteccion de producto para asignar
      producto = Cells(h, colDisponible)
      
'= REVISIÓN POR TASAS DE LAS TIENDAS MENORES A LAS MAYORES
      For cTienda = 1 To numTiendas
   
'==== IDENTIFICACION Y ASIGNACION DE PRODUCTO
         sumaAsigna = 0
      'Identificacion del producto necesario para llegar a siguiente tasa
         If cTienda = numTiendas Then
         'Se fija cuanto asignaria para alcanzar a tasaM si ya estamos asognando a todas las tiendas
            For cMin = 1 To cTienda
               asigna(cMin) = (1.5 * Cells(hMin(numTiendas), colTasa) - Cells(hMin(cMin), colTasa)) / Cells(hMin(cMin), colImpacto)
               sumaAsigna = sumaAsigna + asigna(cMin)
            Next cMin
         Else
         'Tasas menores a la de la posicion cTienda + 1 le alcanzan
            For cMin = 1 To cTienda
               asigna(cMin) = (Cells(hMin(cTienda + 1), colTasa) - Cells(hMin(cMin), colTasa)) / Cells(hMin(cMin), colImpacto)
               sumaAsigna = sumaAsigna + asigna(cMin)
            Next cMin
            tasaAlcanzada = False
         End If
      
      'Asignacion de producto, o todo o la parte correspondiente.
         'Si es posible la asignacion la realiza
         If producto >= sumaAsigna Then
            producto = producto - sumaAsigna
            For cMin = 1 To cTienda
               Cells(hMin(cMin), colAsignado) = asigna(cMin) + Cells(hMin(cMin), colAsignado)
               Cells(hMin(cMin), colTasa) = Cells(hMin(cMin), colTasaIni) + Cells(hMin(cMin), colImpacto) * Cells(hMin(cMin), colAsignado)
            Next cMin
            
            'Si hay producto, continua con el siguiente nivel de asignacion, que se alcance la tasa de la siguiente talla
            If (producto > 0) Then GoTo siguienteNivel
         
         'Si no, entonces reparte lo disponible según lo que le iba a tocar con respecto a lo que se ubiera asignado, así se mantienen tasas semejantes
         Else
            swap = producto
            For cMin = 1 To cTienda
               Cells(hMin(cMin), colAsignado) = producto * asigna(cMin) / sumaAsigna + Cells(hMin(cMin), colAsignado)
               Cells(hMin(cMin), colTasa) = Cells(hMin(cMin), colTasaIni) + Cells(hMin(cMin), colImpacto) * Cells(hMin(cMin), colAsignado)
               swap = swap - producto * asigna(cMin) / sumaAsigna
            Next cMin
            producto = swap
         End If
      
'==== ELABORACION DE RESULTADOS DE ESTA TALLA
      'Redondeo a entero. Primera vez que se resta a lo disponible inicial
         For cMin = 1 To cTienda
            Cells(hMin(cMin), colAsignado) = Int(0.5 + Cells(hMin(cMin), colAsignado))
            Cells(h, colDisponible) = Int(Cells(h, colDisponible) - Cells(hMin(cMin), colAsignado))
            Cells(hMin(cMin), colTasa) = Cells(hMin(cMin), colTasaIni) + Cells(hMin(cMin), colImpacto) * Cells(hMin(cMin), colAsignado)
         Next cMin
         
         If Cells(h, colDisponible) = 0 Then GoTo siguienteTalla 'nada sobro o falto
      
      'Arreglo sobrante por redondeo
         While (Cells(h, colDisponible) > 0)
         'Busca tasa min
            pos = 1
            For cMin = 2 To cTienda
               If Cells(hMin(cMin), colTasa) < Cells(hMin(pos), colTasa) Then
                  pos = cMin
               End If
            Next cMin
         'Asigna un producto
            Cells(hMin(pos), colAsignado) = Cells(hMin(pos), colAsignado) + 1
            Cells(hMin(pos), colTasa) = Cells(hMin(pos), colTasa) + Cells(hMin(pos), colImpacto)
            Cells(h, colDisponible) = Cells(h, colDisponible) - 1
         Wend
         
      'Arreglo faltante por redondeo, quita hasta dejar producto = 0
         While (Cells(h, colDisponible) < 0)
         'Desasigna un producto
            pos = 1
            For cMin = 2 To cTienda
               If (Cells(hMin(cMin), colTasa) > Cells(hMin(pos), colTasa) And Cells(hMin(pos), colAsignado) > 0) Then
                  pos = cMin
               End If
            Next cMin
            Cells(hMin(pos), colAsignado) = Cells(hMin(pos), colAsignado) - 1
            Cells(hMin(pos), colTasa) = Cells(hMin(pos), colTasa) - Cells(hMin(pos), colImpacto)
            Cells(h, colDisponible) = Cells(h, colDisponible) + 1
         Wend

siguienteTalla:
         Exit For

siguienteNivel:
      Next cTienda
      
   'Fila de la siguiente talla
      h = h + numTiendas
      
   'Impresion de avance
      conta = conta + 1
      If conta = 50 Then
         Cells(15, 1) = Cells(15, 1) + avance
         EndTime = Timer
         Cells(16, 1) = Cells(16, 1) + (EndTime - StartTime)
         StartTime = Timer
         Cells(14, 1).Calculate
         Cells(1, 1).Calculate
         Cells(14, 1).Activate
         conta = 0
      End If
      
'Cambio de talla
   Next cTalla

Application.Calculation = xlAutomatic
   Call ActivarFormulaTasa

'Tiempo de ejecucion
   EndTime = Timer
   Cells(16, 1) = Cells(16, 1) + (EndTime - StartTime)
   Cells(14, 1).Activate
   Cells(15, 1) = 1

End Sub

Sub AsignacionSencillaUnoEnUno()
   
   Dim numTallas As Long
   Dim cTalla As Long
   Dim h1 As Long
   Dim h As Long
   
   Dim v As Integer
   Dim pos As Integer
   Dim colTasa As Integer
   Dim colTasaIni As Integer
   Dim colImpacto As Integer
   Dim colAsignado As Integer
   Dim colDisponible As Integer
   
   Dim numTiendas As Integer
   Dim revision As Integer
   Dim producto As Integer
   Dim tasaMin As Integer
   Dim cTienda As Integer
   Dim cMin As Integer
   
   Dim tasaM As Double
   Dim avance As Double
   Dim EndTime As Double
   Dim StartTime As Double
   Dim sumaAsigna As Double
   
   
'Asignacion directa una tienda
   If Cells(24, 1) = 1 Then
      Call CasoUnaTienda
      Exit Sub
   End If


StartTime = Timer
Sheets("Asignacion").Select
   Call DatosIniciales

Application.Calculation = xlManual
   
' Capturar valores
   Cells(15, 1) = 0
   Cells(16, 1) = 0
   tasaM = Cells(20, 1)
   h = 2
   v = 8   ' Inicia en tasa del primero del grupo
   colTasa = v
   colAsignado = v - 1
   colTasaIni = v + 2
   colImpacto = v + 1
   colDisponible = v - 2
   
   numTiendas = Cells(24, 1)
   numTallas = Cells(22, 1)
   revision = numTiendas - 1
   avance = 50 / numTallas
   
   ReDim hMin(numTiendas)
   
   For cTalla = 1 To numTallas
   
   'Primera tasa min
      tasaMin = Cells(h, colTasa)
      pos = 0
      For cMin = 1 To revision
         If tasaMin > Cells(h + cMin, colTasa) Then
            tasaMin = Cells(h + cMin, colTasa)
            pos = cMin
         End If
      Next cMin
   
   'Asigna producto mientras haya y mientras se este por debajo de la tasa deseada
      While (Cells(h, colDisponible) > 0 And tasaM > Cells(h + pos, colTasa))
      'Asigna un producto
         Cells(h + pos, colAsignado) = Cells(h + pos, colAsignado) + 1
         Cells(h + pos, colTasa) = Cells(h + pos, colTasa) + Cells(h + pos, colImpacto)
         Cells(h, colDisponible) = Cells(h, colDisponible) - 1
      ' Busca siguiente tasa min
         tasaMin = Cells(h, colTasa)
         pos = 0
         For cMin = 1 To revision
            If tasaMin > Cells(h + cMin, colTasa) Then
               tasaMin = Cells(h + cMin, colTasa)
               pos = cMin
            End If
         Next cMin
      Wend
      h = h + numTiendas
   
   'Impresion de avance
      conta = conta + 1
      If conta = 50 Then
         Cells(15, 1) = Cells(15, 1) + avance
         EndTime = Timer
         Cells(16, 1) = Cells(16, 1) + (EndTime - StartTime)
         StartTime = Timer
         Cells(14, 1).Calculate
         Cells(1, 1).Calculate
         Cells(14, 1).Activate
         conta = 0
      End If
      
'Cambio de talla
   Next cTalla

Application.Calculation = xlAutomatic
   Call ActivarFormulaTasa

'Tiempo de ejecucion
   EndTime = Timer
   Cells(16, 1) = Cells(16, 1) + (EndTime - StartTime)
   Cells(14, 1).Activate
   Cells(15, 1) = 1
   
End Sub

Sub DatosIniciales()
' Llena los datos de la tasa inicial, impacto y deja vacia la columna de asignacion

'Captura filas para su seleccion
   Dim numTiendas As Integer
   Dim numTallas As Long
   Dim totalFilas As Long
    
Sheets("Asignacion").Select
    
   numTiendas = Cells(24, 1)
   numTallas = Cells(22, 1)
   totalFilas = (numTiendas * numTallas + 1)
    
'Limpia la columna de asignacion
   Range("G2:G" & totalFilas).Select
   Selection.ClearContents

'Crea formula inicial de impacto, tasa inicial y tasa.
   Range("I2").Select
   ActiveCell.FormulaR1C1 = "=IF(RC[-4]>0,100/(RC[-4]),150)"
   Range("J2").Select
   ActiveCell.FormulaR1C1 = "=RC[-1]*(RC[-6])"
   Range("H2").Select
   ActiveCell.FormulaR1C1 = "=RC[2]"
    
'Copia formulas
   Range("H2:J2").Select
   Selection.Copy
    
'Extiende formulas
   Range("H3:J" & totalFilas).Select
   Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
       SkipBlanks:=False, Transpose:=False
   Application.CutCopyMode = False
    
'Convierte a solo valores
   Selection.Copy
   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
       :=False, Transpose:=False
    
   Range("H2").Select
   Application.CutCopyMode = False
   Selection.Copy
   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
       :=False, Transpose:=False
   Application.CutCopyMode = False
End Sub

Sub ActivarFormulaTasa()
    
'Captura filas para su seleccion
    Dim numTiendas As Integer
    Dim numTallas As Long
    Dim totalFilas As Long
    
Sheets("Asignacion").Select
    
    numTiendas = Cells(24, 1)
    numTallas = Cells(22, 1)
    totalFilas = (numTiendas * numTallas + 1)
    
'Pone primera formula
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=RC[2]+RC[1]*RC[-1]"

'La extiende a todas las filas
    Range("H2").Select
    Selection.Copy
    Range("H3:H" & totalFilas).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("H2").Select
    Application.CutCopyMode = False
End Sub

Sub DesignacionPedido()

'Captura filas para su seleccion
   Dim numTiendas As Integer
   Dim numTallas As Long
   Dim totalFilas As Long
    
Sheets("Asignacion").Select
   
   numTiendas = Cells(24, 1)
   
'Si tiendas es uno en la asignación de producto se hizo la designación de pedido
   If numTiendas = 1 Then Exit Sub
   
' Si no, continua
   numTallas = Cells(22, 1)
   totalFilas = (numTiendas * numTallas + 1)

'Crea formulas primera formula
   Range("K2").Select
   ActiveCell.FormulaR1C1 = "=IF( (R28C1 > RC[-3]) , INT((R28C1 - RC[-3])/RC[-2]+0.5) , 0)"
   Range("L2").Select
   ActiveCell.FormulaR1C1 = "=SUM(RC[-1]:R[" & (numTiendas - 1) & "]C[-1])-RC[-6]"

   Range("M2").Select
   ActiveCell.FormulaR1C1 = "=IF( (R30C1 > RC[-5]) , INT((R30C1 - RC[-5])/RC[-4]+0.5) , 0)"
   Range("N2").Select
   ActiveCell.FormulaR1C1 = "=SUM(RC[-1]:R[" & (numTiendas - 1) & "]C[-1])-RC[-8]"

   Range("O2").Select
   ActiveCell.FormulaR1C1 = "=IF( (R32C1 > RC[-7]) , INT((R32C1 - RC[-7])/RC[-6]+0.5) , 0)"
   Range("P2").Select
   ActiveCell.FormulaR1C1 = "=SUM(RC[-1]:R[" & (numTiendas - 1) & "]C[-1])-RC[-10]"

'Extiende formulas talla
   Range("K2").Select
   Selection.Copy
   Range("K3:K" & (numTiendas + 1)).Select
   Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
       SkipBlanks:=False, Transpose:=False
   Range("M2").Select
   Application.CutCopyMode = False
   Selection.Copy
   Range("M3:M" & (numTiendas + 1)).Select
   Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
       SkipBlanks:=False, Transpose:=False
   Range("O2").Select
   Application.CutCopyMode = False
   Selection.Copy
   Range("O3:O" & (numTiendas + 1)).Select
   Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
       SkipBlanks:=False, Transpose:=False
   Range("O5").Select
   Application.CutCopyMode = False
    
'Creacion de clave
  Range("Q2:Q" & (numTiendas + 1)).Select
  Selection = 2
  Range("Q2") = 1

'Selecciona formulas bloque talla y las pega en todas las filas
   If (numTallas > 1) Then
      Range("K2:Q" & (numTiendas + 1)).Select
      Selection.Copy
      Range("K" & (numTiendas + 2) & ":Q" & (totalFilas)).Select
      Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
          SkipBlanks:=False, Transpose:=False
      Selection.End(xlUp).Select
   End If
    
'Limpia hoja de resumen
Sheets("ResumenTalla").Select
   Columns("A:F").Select
   Application.CutCopyMode = False
   Selection.Delete Shift:=xlToLeft
   Range("A1").Select

'Copia columnas para resumen talla
Sheets("Asignacion").Select
   Range("C:C,F:F,L:L,N:N,P:P,Q:Q").Select
   Selection.Copy
    
Sheets("ResumenTalla").Select
   Columns("A:A").Select
   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
       :=False, Transpose:=False
   Application.CutCopyMode = False
    
'Orden por clave
   ActiveWorkbook.Worksheets("ResumenTalla").Sort.SortFields.Clear
   ActiveWorkbook.Worksheets("ResumenTalla").Sort.SortFields.Add Key:=Range( _
       "F2:F" & totalFilas), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
       xlSortNormal
   With ActiveWorkbook.Worksheets("ResumenTalla").Sort
       .SetRange Range("A1:F" & totalFilas)
       .Header = xlYes
       .MatchCase = False
       .Orientation = xlTopToBottom
       .SortMethod = xlPinYin
       .Apply
   End With
   Range("F2").Select
    
'Eliminacion de las filas sobrantes por talla
   If numTiendas > 1 Then
     Rows((numTallas + 2) & ":" & totalFilas).Select
     Selection.Delete Shift:=xlUp
   End If
    
'Ajuste de formato de titulos
   Rows("1:1").Select
   With Selection
       .HorizontalAlignment = xlGeneral
       .VerticalAlignment = xlCenter
       .WrapText = False
       .Orientation = 0
       .AddIndent = False
       .IndentLevel = 0
       .ShrinkToFit = False
       .ReadingOrder = xlContext
       .MergeCells = False
   End With
   With Selection
       .HorizontalAlignment = xlCenter
       .VerticalAlignment = xlCenter
       .WrapText = False
       .Orientation = 0
       .AddIndent = False
       .IndentLevel = 0
       .ShrinkToFit = False
       .ReadingOrder = xlContext
       .MergeCells = False
   End With
   Selection.Font.Bold = True
   With Selection
       .HorizontalAlignment = xlCenter
       .VerticalAlignment = xlCenter
       .WrapText = True
       .Orientation = 0
       .AddIndent = False
       .IndentLevel = 0
       .ShrinkToFit = False
       .ReadingOrder = xlContext
       .MergeCells = False
   End With
    
'Eliminacion de la columna clave
   Columns("F:F").Select
   Selection.Delete Shift:=xlToLeft
    
'Deseleccion
   Range("A1").Select
Sheets("Asignacion").Select
   Range("N2").Select
Sheets("ResumenTalla").Select

End Sub

Sub CasoUnaTienda()
   
   Dim numTallas As Long
   Dim cTalla As Long
   Dim v As Integer
   Dim pos As Integer
   Dim colTasa As Integer
   Dim colTasaIni As Integer
   Dim colImpacto As Integer
   Dim colAsignado As Integer
   Dim colDisponible As Integer
   Dim sumaAsigna As Double
   Dim StartTime As Double
   Dim EndTime As Double
   Dim avance As Double
   Dim tasaM As Double
   
   If Cells(22, 1) = 1 Then
      Call CasoUnaTiendaUnaTalla
      Exit Sub
   End If

StartTime = Timer
Sheets("Asignacion").Select

   
' Capturar valores
   Cells(15, 1) = 0
   Cells(16, 1) = 0
   tasaM = Cells(20, 1)
   v = 8   ' Inicia en tasa del primero del grupo
   colTasa = v
   colAsignado = v - 1
   colTasaIni = v + 2
   colImpacto = v + 1
   colDisponible = v - 2
   
   numTiendas = Cells(24, 1)
   numTallas = Cells(22, 1)
   avance = 50 / numTallas
   
   numTallas = numTallas + 1

Application.Calculation = xlAutomatic

'Valores de impacto y tasa inicial
   'Crea formula inicial de impacto, tasa inicial y tasa.
   Range("I2").Select
   ActiveCell.FormulaR1C1 = "=IF(RC[-4]>0,100/(RC[-4]),150)"
   Range("J2").Select
   ActiveCell.FormulaR1C1 = "=RC[-1]*(RC[-6])"
    
   'Copia formulas
   Range("I2:J2").Select
   Selection.Copy
    
   'Extiende formulas
   Range("I3:J" & numTallas).Select
   Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
       SkipBlanks:=False, Transpose:=False
   Application.CutCopyMode = False
    
   'Convierte a solo valores
   Selection.Copy
   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
       :=False, Transpose:=False

'Asignacion de producto

Application.Calculation = xlManual
   conta = 0
   For cTalla = 2 To numTallas
   
      sumaAsigna = Int(1 + (tasaM - Cells(cTalla, colTasaIni)) / Cells(cTalla, colImpacto))
      If sumaAsigna < 0 Then sumaAsigna = 0
      Cells(cTalla, colDisponible) = Int(0.5 + Cells(cTalla, colDisponible))
   
   'Asignación de lo necesario o lo disponible
      If sumaAsigna <= Cells(cTalla, colDisponible) Then
         Cells(cTalla, colAsignado) = Cells(cTalla, colAsignado) + sumaAsigna
         Cells(cTalla, colDisponible) = Cells(cTalla, colDisponible) - sumaAsigna
      Else
         Cells(cTalla, colAsignado) = Cells(cTalla, colAsignado) + Cells(cTalla, colDisponible)
         Cells(cTalla, colDisponible) = 0
      End If
      
   'Impresion de avance
      conta = conta + 1
      If conta = 50 Then
         Cells(15, 1) = Cells(15, 1) + avance
         EndTime = Timer
         Cells(16, 1) = Cells(16, 1) + (EndTime - StartTime)
         StartTime = Timer
         Cells(14, 1).Calculate
         Cells(1, 1).Calculate
         Cells(14, 1).Activate
         conta = 0
      End If
      
'Cambio de talla
   Next cTalla

   numTallas = numTallas - 1
   
'Cálculo de pedido e ingreso de fórmulas

   Range("H2").Select
   ActiveCell.FormulaR1C1 = "=RC[2]+RC[-1]*RC[+1]"
    
   Range("K2").Select
   ActiveCell.FormulaR1C1 = "=IF( (R28C1 > RC[-3]) , INT((R28C1 - RC[-3])/RC[-2]+0.5) , 0)"
   Range("L2").Select
   ActiveCell.FormulaR1C1 = "=RC[-1]"

   Range("M2").Select
   ActiveCell.FormulaR1C1 = "=IF( (R30C1 > RC[-5]) , INT((R30C1 - RC[-5])/RC[-4]+0.5) , 0)"
   Range("N2").Select
   ActiveCell.FormulaR1C1 = "=RC[-1]"

   Range("O2").Select
   ActiveCell.FormulaR1C1 = "=IF( (R32C1 > RC[-7]) , INT((R32C1 - RC[-7])/RC[-6]+0.5) , 0)"
   Range("P2").Select
   ActiveCell.FormulaR1C1 = "=RC[-1]"

'Extiende formulas talla
   'Tasa
   Range("H2").Select
   Selection.Copy
   Range("H3:H" & (numTallas + 1)).Select
   Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
       SkipBlanks:=False, Transpose:=False
   
   'Pedido
   Range("K2:P2").Select
   Selection.Copy
   Range("K3:P" & (numTallas + 1)).Select
   Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
       SkipBlanks:=False, Transpose:=False
   
'Limpia la hoja de resumen (no necesaria en este caso)
Sheets("ResumenTalla").Select
   Columns("A:F").Select
   Application.CutCopyMode = False
   Selection.Delete Shift:=xlToLeft
   Range("A1").Select

'Ubicación final
Sheets("Asignacion").Select
   Range("G2").Select
   Application.CutCopyMode = False
   Range("G2").Activate
   
Application.Calculation = xlAutomatic

'Tiempo de ejecucion
   EndTime = Timer
   Cells(16, 1) = Cells(16, 1) + (EndTime - StartTime)
   Cells(14, 1).Activate
   Cells(15, 1) = 1
   
End Sub

Sub CasoUnaTiendaTodo()
   
   Dim numTallas As Long
   Dim cTalla As Long
   Dim v As Integer
   Dim pos As Integer
   Dim colTasa As Integer
   Dim colTasaIni As Integer
   Dim colImpacto As Integer
   Dim colAsignado As Integer
   Dim colDisponible As Integer
   Dim sumaAsigna As Double
   Dim StartTime As Double
   Dim EndTime As Double
   Dim avance As Double
   Dim tasaM As Double
   
   If Cells(22, 1) = 1 Then
      Call CasoUnaTiendaUnaTallaTodo
      Exit Sub
   End If


StartTime = Timer
Sheets("Asignacion").Select

   
' Capturar valores
   Cells(15, 1) = 0
   Cells(16, 1) = 0
   tasaM = Cells(20, 1)
   v = 8   ' Inicia en tasa del primero del grupo
   colTasa = v
   colAsignado = v - 1
   colTasaIni = v + 2
   colImpacto = v + 1
   colDisponible = v - 2
   
   numTiendas = Cells(24, 1)
   numTallas = Cells(22, 1)
   avance = 50 / numTallas
   
   numTallas = numTallas + 1

Application.Calculation = xlAutomatic

'Valores de impacto y tasa inicial
   'Crea formula inicial de impacto, tasa inicial y tasa.
   Range("I2").Select
   ActiveCell.FormulaR1C1 = "=IF(RC[-4]>0,100/(RC[-4]),150)"
   Range("J2").Select
   ActiveCell.FormulaR1C1 = "=RC[-1]*(RC[-6])"
    
   'Copia formulas
   Range("I2:J2").Select
   Selection.Copy
    
   'Extiende formulas
   Range("I3:J" & numTallas).Select
   Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
       SkipBlanks:=False, Transpose:=False
   Application.CutCopyMode = False
    
   'Convierte a solo valores
   Selection.Copy
   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
       :=False, Transpose:=False

'Asignacion de producto

Application.Calculation = xlManual
   conta = 0
   For cTalla = 2 To numTallas
   
      sumaAsigna = Int(1 + (tasaM - Cells(cTalla, colTasaIni)) / Cells(cTalla, colImpacto))
      If sumaAsigna < 0 Then sumaAsigna = 0
      Cells(cTalla, colDisponible) = Int(0.5 + Cells(cTalla, colDisponible))
   
   'Asignación de lo necesario o lo disponible
      Cells(cTalla, colAsignado) = Cells(cTalla, colAsignado) + sumaAsigna
      Cells(cTalla, colDisponible) = Cells(cTalla, colDisponible) - sumaAsigna
      
   'Impresion de avance
      conta = conta + 1
      If conta = 50 Then
         Cells(15, 1) = Cells(15, 1) + avance
         EndTime = Timer
         Cells(16, 1) = Cells(16, 1) + (EndTime - StartTime)
         StartTime = Timer
         Cells(14, 1).Calculate
         Cells(1, 1).Calculate
         Cells(14, 1).Activate
         conta = 0
      End If
      
'Cambio de talla
   Next cTalla

   numTallas = numTallas - 1
   
'Cálculo de pedido e ingreso de fórmulas

   Range("H2").Select
   ActiveCell.FormulaR1C1 = "=RC[2]+RC[-1]*RC[+1]"
    
   Range("K2").Select
   ActiveCell.FormulaR1C1 = "=IF( (R28C1 > RC[-3]) , INT((R28C1 - RC[-3])/RC[-2]+0.5) , 0)"
   Range("L2").Select
   ActiveCell.FormulaR1C1 = "=RC[-1]"

   Range("M2").Select
   ActiveCell.FormulaR1C1 = "=IF( (R30C1 > RC[-5]) , INT((R30C1 - RC[-5])/RC[-4]+0.5) , 0)"
   Range("N2").Select
   ActiveCell.FormulaR1C1 = "=RC[-1]"

   Range("O2").Select
   ActiveCell.FormulaR1C1 = "=IF( (R32C1 > RC[-7]) , INT((R32C1 - RC[-7])/RC[-6]+0.5) , 0)"
   Range("P2").Select
   ActiveCell.FormulaR1C1 = "=RC[-1]"

'Extiende formulas talla
   'Tasa
   Range("H2").Select
   Selection.Copy
   Range("H3:H" & (numTallas + 1)).Select
   Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
       SkipBlanks:=False, Transpose:=False
   
   'Pedido
   Range("K2:P2").Select
   Selection.Copy
   Range("K3:P" & (numTallas + 1)).Select
   Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
       SkipBlanks:=False, Transpose:=False
   
'Limpia la hoja de resumen (no necesaria en este caso)
Sheets("ResumenTalla").Select
   Columns("A:F").Select
   Application.CutCopyMode = False
   Selection.Delete Shift:=xlToLeft
   Range("A1").Select

'Ubicación final
Sheets("Asignacion").Select
   Range("G2").Select
   Application.CutCopyMode = False
   Range("G2").Activate
   
Application.Calculation = xlAutomatic

'Tiempo de ejecucion
   EndTime = Timer
   Cells(16, 1) = Cells(16, 1) + (EndTime - StartTime)
   Cells(14, 1).Activate
   Cells(15, 1) = 1
   
End Sub

Sub CasoUnaTiendaUnaTalla()
   
   Dim v As Integer
   Dim pos As Integer
   Dim colTasa As Integer
   Dim colTasaIni As Integer
   Dim colImpacto As Integer
   Dim colAsignado As Integer
   Dim colDisponible As Integer
   Dim sumaAsigna As Double
   Dim StartTime As Double
   Dim EndTime As Double
   Dim tasaM As Double
   

StartTime = Timer
Sheets("Asignacion").Select

   
' Capturar valores
   Cells(15, 1) = 0
   Cells(16, 1) = 0
   tasaM = Cells(20, 1)
   v = 8   ' Inicia en tasa del primero del grupo
   colTasa = v
   colAsignado = v - 1
   colTasaIni = v + 2
   colImpacto = v + 1
   colDisponible = v - 2
   
   numTiendas = Cells(24, 1)

Application.Calculation = xlAutomatic

'Valores de impacto y tasa inicial
   'Crea formula inicial de impacto, tasa inicial y tasa.
   Range("I2").Select
   ActiveCell.FormulaR1C1 = "=IF(RC[-4]>0,100/(RC[-4]),150)"
   Range("J2").Select
   ActiveCell.FormulaR1C1 = "=RC[-1]*(RC[-6])"
    
   'Copia formulas
   Range("I2:J2").Select
   Selection.Copy
    
'Asignacion de producto

Application.Calculation = xlManual
   
   sumaAsigna = Int(1 + (tasaM - Cells(2, colTasaIni)) / Cells(2, colImpacto))
   If sumaAsigna < 0 Then sumaAsigna = 0
   Cells(2, colDisponible) = Int(0.5 + Cells(2, colDisponible))
   
   'Asignación de lo necesario o lo disponible
   If sumaAsigna <= Cells(2, colDisponible) Then
      Cells(2, colAsignado) = Cells(2, colAsignado) + sumaAsigna
      Cells(2, colDisponible) = Cells(2, colDisponible) - sumaAsigna
   Else
      Cells(2, colAsignado) = Cells(2, colAsignado) + Cells(2, colDisponible)
      Cells(2, colDisponible) = 0
   End If
      
'Cálculo de pedido e ingreso de fórmulas

   Range("H2").Select
   ActiveCell.FormulaR1C1 = "=RC[2]+RC[-1]*RC[+1]"
    
   Range("K2").Select
   ActiveCell.FormulaR1C1 = "=IF( (R28C1 > RC[-3]) , INT((R28C1 - RC[-3])/RC[-2]+0.5) , 0)"
   Range("L2").Select
   ActiveCell.FormulaR1C1 = "=RC[-1]"

   Range("M2").Select
   ActiveCell.FormulaR1C1 = "=IF( (R30C1 > RC[-5]) , INT((R30C1 - RC[-5])/RC[-4]+0.5) , 0)"
   Range("N2").Select
   ActiveCell.FormulaR1C1 = "=RC[-1]"

   Range("O2").Select
   ActiveCell.FormulaR1C1 = "=IF( (R32C1 > RC[-7]) , INT((R32C1 - RC[-7])/RC[-6]+0.5) , 0)"
   Range("P2").Select
   ActiveCell.FormulaR1C1 = "=RC[-1]"

'Limpia la hoja de resumen (no necesaria en este caso)
Sheets("ResumenTalla").Select
   Columns("A:F").Select
   Application.CutCopyMode = False
   Selection.Delete Shift:=xlToLeft
   Range("A1").Select

'Ubicación final
Sheets("Asignacion").Select
   Range("G2").Select
   Application.CutCopyMode = False
   Range("G2").Activate
   
Application.Calculation = xlAutomatic

'Tiempo de ejecucion
   EndTime = Timer
   Cells(16, 1) = Cells(16, 1) + (EndTime - StartTime)
   Cells(14, 1).Activate
   Cells(15, 1) = 1
   
End Sub

Sub CasoUnaTiendaUnaTallaTodo()
   
   Dim v As Integer
   Dim pos As Integer
   Dim colTasa As Integer
   Dim colTasaIni As Integer
   Dim colImpacto As Integer
   Dim colAsignado As Integer
   Dim colDisponible As Integer
   Dim sumaAsigna As Double
   Dim StartTime As Double
   Dim EndTime As Double
   Dim tasaM As Double
   

StartTime = Timer
Sheets("Asignacion").Select

   
' Capturar valores
   Cells(15, 1) = 0
   Cells(16, 1) = 0
   tasaM = Cells(20, 1)
   v = 8   ' Inicia en tasa del primero del grupo
   colTasa = v
   colAsignado = v - 1
   colTasaIni = v + 2
   colImpacto = v + 1
   colDisponible = v - 2
   
   numTiendas = Cells(24, 1)

Application.Calculation = xlAutomatic

'Valores de impacto y tasa inicial
   'Crea formula inicial de impacto, tasa inicial y tasa.
   Range("I2").Select
   ActiveCell.FormulaR1C1 = "=IF(RC[-4]>0,100/(RC[-4]),150)"
   Range("J2").Select
   ActiveCell.FormulaR1C1 = "=RC[-1]*(RC[-6])"
    
   'Copia formulas
   Range("I2:J2").Select
   Selection.Copy
    
'Asignacion de producto

Application.Calculation = xlManual
   
   sumaAsigna = Int(1 + (tasaM - Cells(2, colTasaIni)) / Cells(2, colImpacto))
   If sumaAsigna < 0 Then sumaAsigna = 0
   Cells(2, colDisponible) = Int(0.5 + Cells(2, colDisponible))
   
   'Asignación de lo necesario o lo disponible
   Cells(2, colAsignado) = Cells(2, colAsignado) + sumaAsigna
   Cells(2, colDisponible) = Cells(2, colDisponible) - sumaAsigna
      
'Cálculo de pedido e ingreso de fórmulas

   Range("H2").Select
   ActiveCell.FormulaR1C1 = "=RC[2]+RC[-1]*RC[+1]"
    
   Range("K2").Select
   ActiveCell.FormulaR1C1 = "=IF( (R28C1 > RC[-3]) , INT((R28C1 - RC[-3])/RC[-2]+0.5) , 0)"
   Range("L2").Select
   ActiveCell.FormulaR1C1 = "=RC[-1]"

   Range("M2").Select
   ActiveCell.FormulaR1C1 = "=IF( (R30C1 > RC[-5]) , INT((R30C1 - RC[-5])/RC[-4]+0.5) , 0)"
   Range("N2").Select
   ActiveCell.FormulaR1C1 = "=RC[-1]"

   Range("O2").Select
   ActiveCell.FormulaR1C1 = "=IF( (R32C1 > RC[-7]) , INT((R32C1 - RC[-7])/RC[-6]+0.5) , 0)"
   Range("P2").Select
   ActiveCell.FormulaR1C1 = "=RC[-1]"

'Limpia la hoja de resumen (no necesaria en este caso)
Sheets("ResumenTalla").Select
   Columns("A:F").Select
   Application.CutCopyMode = False
   Selection.Delete Shift:=xlToLeft
   Range("A1").Select

'Ubicación final
Sheets("Asignacion").Select
   Range("G2").Select
   Application.CutCopyMode = False
   Range("G2").Activate
   
Application.Calculation = xlAutomatic

'Tiempo de ejecucion
   EndTime = Timer
   Cells(16, 1) = Cells(16, 1) + (EndTime - StartTime)
   Cells(14, 1).Activate
   Cells(15, 1) = 1
   
End Sub
