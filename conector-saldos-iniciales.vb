Option Explicit
'DESARROLLADO POR ING. CHRISTIAN HERRERA CEL: 3154031627
Sub CreaTXT()

    Dim NombreArchivo, RutaArchivo As String
    Dim obj As FileSystemObject
    Dim tx As Scripting.TextStream
    
    Dim Ht As Worksheet
    Dim i, j, nFilas, nColumnas As Integer
    
    NombreArchivo = "Conector Saldos iniciales"
    RutaArchivo = ActiveWorkbook.Path & "\" & NombreArchivo & ".txt"
    
    Set Ht = Worksheets("Movimiento contable")
    
    Set obj = New FileSystemObject
    
    Set tx = obj.CreateTextFile(RutaArchivo)
    
    
    nColumnas = Ht.Range("A1", Ht.Range("A1").End(xlToRight)).Cells.Count
    
    nFilas = Ht.Range("A2", Ht.Range("A2").End(xlDown)).Cells.Count
    'nFilas = Ht.Range("A1", Ht.Range("A1").End(xlDown)).Cells.Count
    
    If (nFilas >= 1000000) Then
    
    nFilas = 1
    
    End If
    
    
    
    
    '###############################################
    
    'creación modulo de saldos inciales
    
    Dim tipo_doc, num_doc, aux_cta_cont, tercero, centro_op_del_movi, unidad_negoci, aux_centro_costos, aux_concep_flujo, valor_debi, valor_credi, valor_base_gravab, tipo_doc_banco, num_doc_banco, obs_movim
    
    
    
    Dim hojaDocContable As Worksheet
    
    Set hojaDocContable = Worksheets("Documento contable")
    
    'desprendemos los campos de esa hoja y los almacenamos en variables
    
    Dim tipoDoc, numDoc, fechaDoc, terceDoc, notasDocCont As String
    
    tipoDoc = hojaDocContable.Range("A2").Value
    
    numDoc = hojaDocContable.Range("B2").Value
    
    fechaDoc = hojaDocContable.Range("C2").Value
    
    terceDoc = hojaDocContable.Range("D2").Value
    
    notasDocCont = hojaDocContable.Range("E2").Value
    
    
    '###############################################
    
    
    'creamos las variables para almacenar los datos, para despues convertirlos en arreglos
    
    Dim reemplazar, cod_item, referencia, ext1, ext2, instalacion, cod_metod, consecuti, referen_comp, extcomp1, extcomp2, cantidad_fab, cantidad_cons, porcentaje, cod_uso, porce_coprod, notas, bodega_cons As String
    
    Dim cadena, cadena2, cadena3, cadena4, cadena5, cadena6, cadena7, cadena8, cadena9, cadena10, cadena11, cadena12, cadena13, cadena14, cadena15, cadena16, cadena17, cadena18 As String
    
    'en cada cadena es donde se van a almacenar los datos
    
    'en este array es donde se van a almacenar los datos, luego se pasan en variables independientes
    
    Dim valores(14) As String
    
    
    
    'empezamos a recorrer las filas y las columnas de excel
    
    tx.WriteLine ("000000100000001001")
    
    'AQUI DEBEMOS LLAMAR A LA HOJA #1 LA CUAL ES: DOCUMENTO CONTABLE PARA OBTENER LOS DATOS Y DISPONERLOS SOBRE LA SEGUNDA FILA
    
    
    'ahora debemos hacer un separador para que entre la linea de notas y todo nos de 255 de longitud
    
    Dim separador As String
    
    separador = 255 - Len(notasDocCont)
    
    'creamos el separador entre el tercero del documento y las notas
    
    Dim separador_2 As String
    
    separador_2 = 15 - Len(terceDoc)
                
    tx.WriteLine ("0000002" + "035000010011001" & tipoDoc & numDoc & fechaDoc & terceDoc & Space(separador_2) & "0003000" & notasDocCont + Space(separador))
                
    '###############
    
    Dim k As Integer
    
    k = 2
    
    For i = 1 To nFilas
    
        For j = 0 To nColumnas
        
        valores(j) = ActiveCell.Value
        
        Cells(i + 1, j + 1).Select
        
        'primer valor de la primera cadena
        
        'cadena = cadena + valores(j)
        
        cadena = valores(1)

        cadena2 = valores(2)
            
        cadena3 = valores(3)
            
        cadena4 = valores(4)
            
        cadena5 = valores(5)
            
        cadena6 = valores(6)
    
        cadena7 = valores(7)
        
        cadena8 = valores(8)
        
        cadena9 = valores(9)
        
        cadena10 = valores(10)
        
        cadena11 = valores(11)
        
        cadena12 = valores(12)
        
        cadena13 = valores(13)
        
        cadena14 = valores(14)
        
        
        
        Next j
        
        Dim ite As String
        
        ite = i + 2
        
        'creamos el primer separador para que no se pase de 7 caracteres de longitud
        
        Dim primer_separador, primer_s, result_primer_sep As String
        
        primer_separador = 7
    
        primer_s = Len(ite)
        
        result_primer_sep = primer_separador - primer_s
        
        
        'SEPARADOR ENTRE AUXILIAR DE CUENTA CONTABLE Y TERCERO, SON 12 ESPACIOS
        
        
        Dim p_s, primer_sep, result_primer_s As String
        
        p_s = 15
        primer_sep = Len(cadena4)
        
        result_primer_s = p_s - primer_sep
        
        
        'SEPARADOR ENTRE UNIDAD DE NEGOCIO Y AUXILIAR DE CENTRO DE COSTOS
        
        
        Dim s_s, segundo_sep, result_segundo_s As String
        
        s_s = 22
        segundo_sep = Len(cadena7)
        
        result_segundo_s = s_s - segundo_sep
        
        
        'SEPARADOR ENTRE AUXILIAR DE CENTRO DE COSTOS Y AUXILIAR DE CONCEPTO DE FLUJO EFECTIVO
        
        Dim t_s, tercer_sep, result_tercer_s As String
        
        t_s = 15
        tercer_sep = Len(cadena8)
        
        result_tercer_s = t_s - tercer_sep
    
    
        '##############################################################################
        
        'AHORA VIENE LA PARTE DONDE EMPEZAMOS A JUGAR CON LOS ENTEROS Y LOS DECIMALES
        
        Dim posicion As Integer
        'Dim symbol As String

        posicion = InStr(cadena9, ".")
        'symbol = Mid(cadena9, posicion, 1)
        
        Dim primer_cero, segundo_cero, izq, dere As String

        'MsgBox symbol
        
        Dim punto As String
        
        'creamos una condición por si hay numeros que son enteros o sea no encontró el .
        
        If (posicion = 0) Then
        
        primer_cero = 15 - Len(cadena9)
        
        segundo_cero = 4
        
        punto = "."
        
        Else
        
        
        'ahora vamos a extraer lo que contiene el . antes
        
        Dim izquierda As String

        izquierda = Left(cadena9, posicion - 1)

        'MsgBox izquierda
        
        'ahora vamos a extraer lo que contiene el . despues
        
        Dim derecha As String

        derecha = Mid(cadena9, posicion + 1)

        'MsgBox derecha
        
        'ahora vamos a agregar los datos para que no se pasen de ciertos caracteres
        
        izq = Len(izquierda)
        
        dere = Len(derecha)
        
        
        primer_cero = 15 - izq
        segundo_cero = 4 - dere
        
        
        End If
        
        '###############################################################
        
        'SEGUNDO JUEGUE CON LOS ENTEROS Y DECIMALES
        
        Dim posicion_dos As Integer
        'Dim symbol As String

        posicion_dos = InStr(cadena10, ".")
        'symbol = Mid(cadena9, posicion, 1)
        
        Dim primer_cero_primero, segundo_cero_segundo, izq_prim, dere_segun As String

        'MsgBox symbol
        
        Dim punto_segun As String
        
        'creamos una condición por si hay numeros que son enteros o sea no encontró el .
        
        If (posicion_dos = 0) Then
        
        primer_cero_primero = 15 - Len(cadena10)
        
        segundo_cero_segundo = 4
        
        punto_segun = "."
        
        Else
        
        
        'ahora vamos a extraer lo que contiene el . antes
        
        Dim izquierda_prim As String

        izquierda_prim = Left(cadena10, posicion_dos - 1)

        'MsgBox izquierda
        
        'ahora vamos a extraer lo que contiene el . despues
        
        Dim derecha_segun As String

        derecha_segun = Mid(cadena10, posicion_dos + 1)

        'MsgBox derecha
        
        'ahora vamos a agregar los datos para que no se pasen de ciertos caracteres
        
        izq_prim = Len(izquierda_prim)
        
        dere_segun = Len(derecha_segun)
        
        
        primer_cero_primero = 15 - izq_prim
        segundo_cero_segundo = 4 - dere_segun
        
        punto_segun = ""
        
        End If
        
        '###############################################################
        
        'TERCER JUEGUE CON LOS ENTEROS Y DECIMALES (VALOR BASE GRABABLE)
        
        Dim posicion_tres As Integer
        'Dim symbol As String

        posicion_tres = InStr(cadena11, ".")
        'symbol = Mid(cadena9, posicion, 1)
        
        Dim primer_cero_tercero, segundo_cero_tercero, izq_tercero, dere_tercero As String

        'MsgBox symbol
        
        Dim punto_tercer As String
        
        'creamos una condición por si hay numeros que son enteros o sea no encontró el .
        
        If (posicion_tres = 0) Then
        
        primer_cero_tercero = 15 - Len(cadena11)
        
        segundo_cero_tercero = 4
        
        punto_tercer = "."
        
        Else
        
        
        'ahora vamos a extraer lo que contiene el . antes
        
        Dim izquierda_tercer As String

        izquierda_tercer = Left(cadena11, posicion_tres - 1)

        'MsgBox izquierda
        
        'ahora vamos a extraer lo que contiene el . despues
        
        Dim derecha_tercer As String

        derecha_tercer = Mid(cadena11, posicion_tres + 1)

        'MsgBox derecha
        
        'ahora vamos a agregar los datos para que no se pasen de ciertos caracteres
        
        izq_tercero = Len(izquierda_tercer)
        
        dere_tercero = Len(derecha_tercer)
        
        
        primer_cero_tercero = 15 - izq_tercero
        segundo_cero_tercero = 4 - dere_tercero
        
        punto_tercer = ""
        
        End If
        
        '######
        
        'agregamos los ceros que viene en el txt
        
        Dim ceros As String
        
        ceros = "+000000000000000.0000+000000000000000.0000+"
        
        'ahora hacemos un comprobación para que cuando el valor del tipó de documento de banco esté en 0 se coloquen 2 espacios
        
        'comprobación si la cadena13 (tipo de documento de banco) está vacio agregue 2 espacios en blanco
        
        If (cadena12 = Empty) Then
            
            cadena12 = "  "
            
        
        End If
        
        
        'separador para que el campo de numero de documento de banco no supere los 7 caracteres
        
        
        Dim segundo_separador, segundo_s, result_segundo_sep As String
        
        segundo_separador = 8
    
        segundo_s = Len(cadena13)
        
        result_segundo_sep = segundo_separador - segundo_s
        
        
        '###############################################################
        
         
        'tx.WriteLine (String(result_primer_sep, "0") + ite + ("0820000500100000000") + cadena + cadena2 + cadena3 + Space(result_primer_s) + cadena4 + Space(result_segundo_s) + cadena5 + Space(result_tercer_s) + cadena6 + cadena7 + String(result_segundo_sep, "0") + cadena8 + "0000000" + cadena9 + Space(result_cuarto_s) + cadena10 + Space(result_quinto_s) + cadena11 + Space(result_sexto_s) + String(result_tercer_sep, "0") + cadena12 + "." + "0000" + String(result_cuarto_sep, "0") + cadena13 + "." + "000000" + "20220101" + Space(8) + String(primer_cero, "0") + cadena14 + punto + String(segundo_cero, "0") + "0000" + cadena15 + String(primer_cero_dos, "0") + cadena16 + punto_dos + String(segundo_cero_dos, "0") + cadena17 + Space(result_septimo_s) + cadena18 + Space(10))
        
        'tx.WriteLine (String(result_primer_sep, "0") + ite + ("08200005001") + cadena + ("000000") + cadena + cadena2 + cadena3 + Space(result_primer_s) + cadena4 + Space(result_segundo_s) + cadena5 + Space(result_tercer_s) + cadena6 + cadena7 + String(result_segundo_sep, "0") + cadena8 + "0000000" + cadena9 + Space(result_cuarto_s) + cadena10 + Space(result_quinto_s) + cadena11 + Space(result_sexto_s) + String(primer_cero_tres, "0") + cadena12 + punto_tres + String(segundo_cero_tres, "0") + String(primer_cero_cuatro, "0") + cadena13 + punto_cuatro + String(segundo_cero_cuatro, "0") + "20220101" + Space(8) + String(primer_cero, "0") + cadena14 + punto + String(segundo_cero, "0") + "0000" + cadena15 + String(primer_cero_dos, "0") + cadena16 + punto_dos + String(segundo_cero_dos, "0") + cadena17 + Space(result_septimo_s) + cadena18 + "0" + Space(10))
        
        tx.WriteLine (String(result_primer_sep, "0") + ite + "03510002001001" + cadena + cadena2 + cadena3 + Space(12) + cadena4 + Space(result_primer_s) + cadena5 + cadena6 + Space(result_segundo_s) + cadena7 + Space(result_tercer_s) + cadena8 + Space(6) + "+" + String(primer_cero, "0") + cadena9 + punto + String(segundo_cero, "0") + "+" + String(primer_cero_primero, "0") + cadena10 + punto_segun + String(segundo_cero_segundo, "0") + ceros + String(primer_cero_tercero, "0") + cadena11 + punto_tercer + String(segundo_cero_tercero, "0") + cadena12 + String(result_segundo_sep, "0") + cadena13 + cadena14 + Space(240))
        

            
    Next i
    
    
    'aqui debemos hacer la comprobacion si el mov de cxc está diligenciado para así ejecutar la función de lo contrario no se ejecuta
    
    Dim hojaMovCxc As Worksheet
    
    Set hojaMovCxc = Worksheets("Movimiento CxC")
    
    'desprendemos los campos de esa hoja y los almacenamos en variables, en este caso solo vamos a validar que el campo detipo de doc, esté diligenciado
    
    Dim tipoDocHoja
    
    tipoDocHoja = hojaMovCxc.Range("A2").Value
    
    
    If (tipoDocHoja <> Empty) Then
    
    Call MovCxC(ite, tx, result_primer_sep, hojaMovCxc)
    
    End If
    
    'AHORA POR ULTIMO HACEMOS LO MISMO DE LA ANTIGUA HOJA PERO CON MOVIMIENTO CXP
    
    Dim hojaMovCxp As Worksheet
    
    Set hojaMovCxp = Worksheets("Movimiento CxP")
    
    Dim tipoDocHojaCxp
    
    tipoDocHojaCxp = hojaMovCxp.Range("A2").Value
    
    If (tipoDocHojaCxp <> Empty) Then
    
    Call MovCxP(tx, result_primer_sep, hojaMovCxp, result_segundo_sep)
    
    End If
    
    
    '#######################################################################################################
    
    Dim ite_4 As String
    
    ite_4 = Range("AG1").Value
    
    Dim ultimo_separador, ultimo_s, result_ultimo_sep As String
    
    ultimo_separador = 7
        
    ultimo_s = Len(ite_4)
        
    result_ultimo_sep = ultimo_separador - ultimo_s
    
    'tx.WriteLine (String(result_ultimo_sep, "0") + ite + 1)
    
    ite = ite + 1
    
    If (ite_4 = "") Then
    
    ite_4 = ite
    
    ultimo_s = Len(ite)
    
    result_ultimo_sep = ultimo_separador - ultimo_s
    
    End If
    
    tx.WriteLine (String(result_ultimo_sep, "0") + ite_4 + ("99990001001"))
          
    
    
    tx.Close
    
    'Set obj = Nothing
    
    MsgBox "El archivo plano se ha generado con exito..."
    
    Sheets("Movimiento contable").Select

End Sub

'#######################################################################################################
        
'AHORA TENEMOS QUE TRAER LOS ELEMENTOS DE LA PAGINA DE "MOVIMIENTO CXC" Y DESDE AHI TRAER TODOS LOS DEMAS ELEMENTOS, DEBEMOS HACER LO MISMO QUE CUANDO NOS TRAJIMOS LOS DATOS MEDIANTE UN BUCLE FOR
    

Sub MovCxC(ByVal ite, ByVal tx, ByVal result_primer_sep, ByVal hojaMovCxc)

    'seleccionamos la hoja en la cual nos encontramos
    
    Sheets("Movimiento CxC").Select

    Dim ite_2 As String
    
    
    ite_2 = ite + 1
    
    Dim nFilasMovCxc, nColumnasMovCxc As Integer
    

    'nFilasMovCxc = hojaMovCxc.Range("A2", hojaMovCxc.Range("A2").End(xlDown)).Cells.Count
    nFilasMovCxc = hojaMovCxc.Range("A2", hojaMovCxc.Range("A2").End(xlDown)).Cells.Count
    
    nColumnasMovCxc = hojaMovCxc.Range("A1", hojaMovCxc.Range("A1").End(xlToRight)).Cells.Count
    
    Dim cadena, cadena2, cadena3, cadena4, cadena5, cadena6, cadena7, cadena8, cadena9, cadena10, cadena11, cadena12, cadena13, cadena14, cadena15, cadena16, cadena17 As String
    
    'en cada cadena es donde se van a almacenar los datos
    
    'en este array es donde se van a almacenar los datos, luego se pasan en variables independientes
    
    Dim valores(17) As String
    
    If (nFilasMovCxc >= 100000) Then
    
    nFilasMovCxc = 1
    
    End If
    
    
    Dim k, i, j As Integer
    
    k = 2
    
    For i = 1 To nFilasMovCxc
    
        For j = 0 To nColumnasMovCxc
        
        valores(j) = ActiveCell.Value
        
        Cells(i + 1, j + 1).Select
        
        'primer valor de la primera cadena
        
        'cadena = cadena + valores(j)
        
        cadena = valores(1)

        cadena2 = valores(2)
            
        cadena3 = valores(3)
            
        cadena4 = valores(4)
            
        cadena5 = valores(5)
            
        cadena6 = valores(6)
    
        cadena7 = valores(7)
        
        cadena8 = valores(8)
        
        cadena9 = valores(9)
        
        cadena10 = valores(10)
        
        cadena11 = valores(11)
        
        cadena12 = valores(12)
        
        cadena13 = valores(13)
        
        cadena14 = valores(14)
        
        cadena15 = valores(15)
        
        cadena16 = valores(16)
        
        cadena17 = valores(17)
        
        
        Next j
        
        
        'primer separador para el CXC
        
        Dim primero_separador, primero_s, result_primero_sep As String
        
        primero_separador = 7
    
        primero_s = Len(ite_2)
        
        result_primero_sep = primero_separador - primero_s
        
        
        'creamos el segundo separador para que no se pase de 15 caracteres de longitud
        
        Dim segundo_separador, segundo_s, result_segundo_sep As String
        
        segundo_separador = 15
    
        segundo_s = Len(cadena4)
        
        result_segundo_sep = segundo_separador - segundo_s
        
        
        'creamos el tercer separador para que no se pase de 38 caracteres de longitud espacio entre valor debito y valor credito
        
        Dim tercer_separador, tercer_s, result_tercer_sep As String
        
        tercer_separador = 36
    
        tercer_s = Len(cadena5)
        
        result_tercer_sep = tercer_separador - tercer_s
        
        '##############################################################################
        
        'AHORA VIENE LA PARTE DONDE EMPEZAMOS A JUGAR CON LOS ENTEROS Y LOS DECIMALES
        
        Dim posicion As Integer
        'Dim symbol As String

        posicion = InStr(cadena7, ".")
        'symbol = Mid(cadena9, posicion, 1)
        
        Dim primer_cero, segundo_cero, izq, dere As String

        'MsgBox symbol
        
        Dim punto As String
        
        'creamos una condición por si hay numeros que son enteros o sea no encontró el .
        
        If (posicion = 0) Then
        
        primer_cero = 15 - Len(cadena7)
        
        segundo_cero = 4
        
        punto = "."
        
        Else
        
        
        'ahora vamos a extraer lo que contiene el . antes
        
        Dim izquierda As String

        izquierda = Left(cadena7, posicion - 1)

        'MsgBox izquierda
        
        'ahora vamos a extraer lo que contiene el . despues
        
        Dim derecha As String

        derecha = Mid(cadena7, posicion + 1)

        'MsgBox derecha
        
        'ahora vamos a agregar los datos para que no se pasen de ciertos caracteres
        
        izq = Len(izquierda)
        
        dere = Len(derecha)
        
        
        primer_cero = 15 - izq
        segundo_cero = 4 - dere
        
        
        End If
        
        '###############################################################
        
        
        '###############################################################
        
        'SEGUNDO JUEGUE CON LOS ENTEROS Y DECIMALES
        
        Dim posicion_dos As Integer
        'Dim symbol As String

        posicion_dos = InStr(cadena8, ".")
        'symbol = Mid(cadena9, posicion, 1)
        
        Dim primer_cero_primero, segundo_cero_segundo, izq_prim, dere_segun As String

        'MsgBox symbol
        
        Dim punto_segun As String
        
        'creamos una condición por si hay numeros que son enteros o sea no encontró el .
        
        If (posicion_dos = 0) Then
        
        primer_cero_primero = 15 - Len(cadena8)
        
        segundo_cero_segundo = 4
        
        punto_segun = "."
        
        Else
        
        
        'ahora vamos a extraer lo que contiene el . antes
        
        Dim izquierda_prim As String

        izquierda_prim = Left(cadena8, posicion_dos - 1)

        'MsgBox izquierda
        
        'ahora vamos a extraer lo que contiene el . despues
        
        Dim derecha_segun As String

        derecha_segun = Mid(cadena8, posicion_dos + 1)

        'MsgBox derecha
        
        'ahora vamos a agregar los datos para que no se pasen de ciertos caracteres
        
        izq_prim = Len(izquierda_prim)
        
        dere_segun = Len(derecha_segun)
        
        
        primer_cero_primero = 15 - izq_prim
        segundo_cero_segundo = 4 - dere_segun
        
        punto_segun = ""
        
        End If
        
        '###############################################################
        
        
        '###############################################################
        

        
        Dim ceros As String
        
        ceros = "+000000000000000.0000+000000000000000.0000"
        
        
        '##################################################################
        
        'cuarto separador entre observacion del movimiento y la sucursal del cliente
        
        Dim cuarto_separador, cuarto_s, result_cuarto_sep As String
        
        cuarto_separador = 255
    
        cuarto_s = Len(cadena9)
        
        result_cuarto_sep = cuarto_separador - cuarto_s
        
        
        
        '###############################################################################
        
        'rellenador de ceros para la columna 12 en cxc
        
        Dim doce_separador, doce_s, result_doce_sep As String
        
        doce_separador = 8
    
        doce_s = Len(cadena12)
        
        result_doce_sep = doce_separador - doce_s
        
        
        '#####################################################################
        
        'el ite_3 es para que aumente de 1 en 1
            
        'Dim ite_3 As Integer
        
        'ite_3 = -1
        
        'creamos el quinto separador para que no se pase de 3 caracteres de longitud
        
        Dim quinto_separador, quinto_s, result_quinto_sep As String
        
        quinto_separador = 3
    
        quinto_s = Len(cadena13)
        
        result_quinto_sep = quinto_separador - quinto_s
        
        Dim ceros_dos As String
        
        ceros_dos = "+000000000000000.0000+000000000000000.0000+000000000000000.0000+000000000000000.0000+000000000000000.0000+000000000000000.0000+000000000000000.0000"
        
        
        Dim separador_3 As String
        
        separador_3 = 255 - Len(cadena17)
        
        tx.WriteLine (String(result_primero_sep, "0") + ite_2 + "03510102001001" + cadena + cadena2 + cadena3 + Space(12) + cadena4 + Space(result_segundo_sep) + cadena5 + cadena6 + Space(result_tercer_sep) + "+" + String(primer_cero, "0") + cadena7 + punto + String(segundo_cero, "0") + "+" + String(primer_cero_primero, "0") + cadena8 + punto_segun + String(segundo_cero_segundo, "0") + ceros + cadena9 + Space(result_cuarto_sep) + cadena10 + cadena11 + String(result_doce_sep, "0") + cadena12 + String(result_quinto_sep, "0") & cadena13 + cadena14 + cadena15 + ceros_dos + cadena16 + Space(6) + cadena17 + Space(separador_3))
        
        ite_2 = ite_2 + 1
        
        
        'vamos a asignar la variable de ite_2 a una celda para luego recuperarla ya que vba no permite instanciar una variable de un sub o procedimiento
        
        'la vamos a almacenar en la celda "AG1"
        
        Range("AG1").Value = ite_2
        
        
        'tx.WriteLine (String(result_primer_sep, "0") + ite + "03510002001001" + cadena + cadena2 + cadena3 + Space(12) + cadena4 + Space(result_primer_s) + cadena5 + cadena6 + Space(result_segundo_s) + cadena7 + Space(result_tercer_s) + cadena8 + Space(6) + "+" + String(primer_cero, "0") + cadena9 + punto + String(segundo_cero, "0") + "+" + String(primer_cero_primero, "0") + cadena10 + punto_segun + String(segundo_cero_segundo, "0") + ceros + String(primer_cero_tercero, "0") + cadena11 + punto_tercer + String(segundo_cero_tercero, "0") + cadena12 + String(result_segundo_sep, "0") + cadena13 + cadena14 + Space(240))
        
    Next i
    
    
    
    'MsgBox ("hay" & nFilasMovCxc & "filas" & nColumnasMovCxc & "columnas")
    

    'MsgBox ("hay dat   os" & ite)
    
    
    
    'tx.WriteLine (ite)
    
End Sub
Sub MovCxP(ByVal tx, ByVal result_primer_sep, ByVal hojaMovCxp, ByVal result_segundo_sep)


    'vamos a traer el valor del ite que se encuentra almacenado en la celda "AG1" del libro mov CxC

    Dim ite_3 As String

    ite_3 = Range("AG1").Value
    
    

    Sheets("Movimiento CxP").Select
    
    Dim nFilasMovCxp, nColumnasMovCxp As Integer
    

    nFilasMovCxp = hojaMovCxp.Range("A2", hojaMovCxp.Range("A2").End(xlDown)).Cells.Count
    nColumnasMovCxp = hojaMovCxp.Range("A1", hojaMovCxp.Range("A1").End(xlToRight)).Cells.Count
    
    
    If (nFilasMovCxp >= 1000000) Then
    
    nFilasMovCxp = 1
    
    End If
    
    
    
    Dim cadena, cadena2, cadena3, cadena4, cadena5, cadena6, cadena7, cadena8, cadena9, cadena10, cadena11, cadena12, cadena13, cadena14, cadena15, cadena16, cadena17 As String
    
    'en cada cadena es donde se van a almacenar los datos
    
    'en este array es donde se van a almacenar los datos, luego se pasan en variables independientes
    
    Dim valores(17) As String
    
    
    Dim i, j As Integer
    
    For i = 1 To nFilasMovCxp
    
        For j = 0 To nColumnasMovCxp
        
        valores(j) = ActiveCell.Value
        
        Cells(i + 1, j + 1).Select
        
        'primer valor de la primera cadena
        
        'cadena = cadena + valores(j)
        
        cadena = valores(1)

        cadena2 = valores(2)
            
        cadena3 = valores(3)
            
        cadena4 = valores(4)
            
        cadena5 = valores(5)
            
        cadena6 = valores(6)
    
        cadena7 = valores(7)
        
        cadena8 = valores(8)
        
        cadena9 = valores(9)
        
        cadena10 = valores(10)
        
        cadena11 = valores(11)
        
        cadena12 = valores(12)
        
        cadena13 = valores(13)
        
        cadena14 = valores(14)
        
        cadena15 = valores(15)
        
        cadena16 = valores(16)
        
        cadena17 = valores(17)
        
        
        Next j
        
        
        'creamos el tercer separador para que no se pase de 38 caracteres de longitud espacio entre valor debito y valor credito
        
        Dim tercer_separador, tercer_s, result_tercer_sep As String
        
        tercer_separador = 36
    
        tercer_s = Len(cadena5)
        
        result_tercer_sep = tercer_separador - tercer_s
        
        
        '##############################################################################
        
        'AHORA VIENE LA PARTE DONDE EMPEZAMOS A JUGAR CON LOS ENTEROS Y LOS DECIMALES
        
        Dim posicion As Integer
        'Dim symbol As String

        posicion = InStr(cadena7, ".")
        'symbol = Mid(cadena9, posicion, 1)
        
        Dim primer_cero, segundo_cero, izq, dere As String

        'MsgBox symbol
        
        Dim punto As String
        
        'creamos una condición por si hay numeros que son enteros o sea no encontró el .
        
        If (posicion = 0) Then
        
        primer_cero = 15 - Len(cadena7)
        
        segundo_cero = 4
        
        punto = "."
        
        Else
        
        
        'ahora vamos a extraer lo que contiene el . antes
        
        Dim izquierda As String

        izquierda = Left(cadena7, posicion - 1)

        'MsgBox izquierda
        
        'ahora vamos a extraer lo que contiene el . despues
        
        Dim derecha As String

        derecha = Mid(cadena7, posicion + 1)

        'MsgBox derecha
        
        'ahora vamos a agregar los datos para que no se pasen de ciertos caracteres
        
        izq = Len(izquierda)
        
        dere = Len(derecha)
        
        
        primer_cero = 15 - izq
        segundo_cero = 4 - dere
        
        
        End If
        
        '###############################################################
        
        
        '###############################################################
        
        'SEGUNDO JUEGUE CON LOS ENTEROS Y DECIMALES
        
        Dim posicion_dos As Integer
        'Dim symbol As String

        posicion_dos = InStr(cadena8, ".")
        'symbol = Mid(cadena9, posicion, 1)
        
        Dim primer_cero_primero, segundo_cero_segundo, izq_prim, dere_segun As String

        'MsgBox symbol
        
        Dim punto_segun As String
        
        'creamos una condición por si hay numeros que son enteros o sea no encontró el .
        
        If (posicion_dos = 0) Then
        
        primer_cero_primero = 15 - Len(cadena8)
        
        segundo_cero_segundo = 4
        
        punto_segun = "."
        
        Else
        
        
        'ahora vamos a extraer lo que contiene el . antes
        
        Dim izquierda_prim As String

        izquierda_prim = Left(cadena8, posicion_dos - 1)

        'MsgBox izquierda
        
        'ahora vamos a extraer lo que contiene el . despues
        
        Dim derecha_segun As String

        derecha_segun = Mid(cadena8, posicion_dos + 1)

        'MsgBox derecha
        
        'ahora vamos a agregar los datos para que no se pasen de ciertos caracteres
        
        izq_prim = Len(izquierda_prim)
        
        dere_segun = Len(derecha_segun)
        
        
        primer_cero_primero = 15 - izq_prim
        segundo_cero_segundo = 4 - dere_segun
        
        punto_segun = ""
        
        End If
        
        '###############################################################
        
        'agregamos los ceros que viene en el txt
        
        Dim ceros As String
        
        ceros = "+000000000000000.0000+000000000000000.0000"
        
        
        '##################################################################
        
        'cuarto separador entre observacion del movimiento y la sucursal del cliente
        
        Dim cuarto_separador, cuarto_s, result_cuarto_sep As String
        
        cuarto_separador = 255
    
        cuarto_s = Len(cadena9)
        
        result_cuarto_sep = cuarto_separador - cuarto_s
        
        
        '##################################################################
        
        'quinto separador entre el prefijo de numero de cruce y el numero de documento de cruce
        
        Dim quinto_separador, quinto_s, result_quinto_sep As String
        
        quinto_separador = 20
    
        quinto_s = Len(cadena11)
        
        result_quinto_sep = quinto_separador - quinto_s
        
        
        '##################################################################
        
        'sexto separador para no pasarnos de la longitud de 3 caracteres para el caso del numero de cuota de documento de cruce
        
        Dim sexto_separador, sexto_s, result_sexto_sep As String
        
        sexto_separador = 3
    
        sexto_s = Len(cadena13)
        
        result_sexto_sep = sexto_separador - sexto_s
        
        
        
        Dim ceros_tres As String
        
        ceros_tres = "+000000000000000.0000+000000000000000.0000+000000000000000.0000+000000000000000.0000+000000000000000.0000"
        
        
        'creamos el quinto separador para que no se pase de 3 caracteres de longitud
        
        Dim septimo_separador, septimo_s, result_septimo_sep As String
        
        septimo_separador = 255
    
        septimo_s = Len(cadena17)
        
        result_septimo_sep = septimo_separador - septimo_s
        
        'tx.WriteLine (String(result_primer_sep, "0") + ite_2 + "03510102001001" + cadena + cadena2 + cadena3 + Space(12) + cadena4 + Space(result_segundo_sep) + cadena5 + cadena6 + Space(result_tercer_sep) + "+" + String(primer_cero, "0") + cadena7 + punto + String(segundo_cero, "0") + "+" + String(primer_cero_primero, "0") + cadena8 + punto_segun + String(segundo_cero_segundo, "0") + ceros + cadena9 + Space(result_cuarto_sep) + cadena10 + cadena11 + cadena12 + String(result_quinto_sep, "0") & cadena13 + cadena14 + cadena15 + ceros_dos + cadena16 + Space(6) + cadena17)
        
        tx.WriteLine (String(result_primer_sep, "0") + ite_3 + "03510203001001" + cadena + cadena2 + cadena3 + Space(12) + cadena4 + Space(result_segundo_sep - 1) + cadena5 + cadena6 + Space(result_tercer_sep) + "+" + String(primer_cero, "0") + cadena7 + punto + String(segundo_cero, "0") + "+" + String(primer_cero_primero, "0") + cadena8 + punto_segun + String(segundo_cero_segundo, "0") + ceros + cadena9 + Space(result_cuarto_sep) + cadena10 + cadena11 + Space(result_quinto_sep) + cadena12 + String(result_sexto_sep, "0") + cadena13 + Space(10) + cadena14 + cadena15 + cadena16 + ceros_tres + cadena17 + Space(result_septimo_sep))
        
        ite_3 = ite_3 + 1
        
        'agregamos el ite_3 para dejarlo en la hoja
        
        Range("AG1").Value = ite_3
        
Next i

End Sub
