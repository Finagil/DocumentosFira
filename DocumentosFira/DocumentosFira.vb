Imports System.IO
Imports System.Net.Mail
Imports System.Data.SqlClient



Module DocumentosFira
    'Dim Rutafira As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.GetName.CodeBase)
    Dim porcentaje As Double
    Dim tipoCre As String
    Dim total As Double
    Dim RutaFira As String = My.Settings.RutaFira
    Dim RutaFiraMes As String
    Dim RutaPol As String = My.Settings.RutaPolizas
    Dim Errores As Boolean
    Dim ErrorControl As New EventLog
    Dim ConnStr As String = "Data Source=SERVER-RAID2;Initial Catalog=production;Persist Security Info=True;User ID=User_PRO;Password=User_PRO2015"
    Dim Conn As New SqlClient.SqlConnection(ConnStr)
    Dim Conn2 As New SqlClient.SqlConnection(ConnStr)
    Dim Conn3 As New SqlClient.SqlConnection(ConnStr)
    Dim Conn4 As New SqlClient.SqlConnection(ConnStr)
    Dim fecha As Date

    Sub Main()
        'Console.WriteLine("Saldos Insolutos")
        'SaldosInsolutos()
        Console.WriteLine("Antigúedad de Saldos")
        Antiguedad()
        Console.WriteLine("Guarda Historia")
        LlenaHisotriaCartera()
        fecha = Date.Now
        fecha = fecha.AddDays(-8)
        RutaFiraMes = RutaFira & "\" & MonthName(fecha.Month)
        Errores = False
        ErrorControl = New EventLog("Application", System.Net.Dns.GetHostName(), "Docuemntos Fira")
        MantieneCarpetas()
        InsertaDatosPasivoFira()
        Console.WriteLine("Fira Leé Archivos Tipo 3")
        LeeCSV_Tipo3()
        Call RevisaArchivos()
        If Errores = False Or Errores = True Then
            Dim di1 As New DirectoryInfo(RutaPol)
            Dim f1 As FileInfo() = di1.GetFiles("*.pol")
            Dim queryBORRAR As New SqlClient.SqlCommand
            If Conn.State <> ConnectionState.Open Then Conn.Open()
            queryBORRAR.Connection = Conn
            queryBORRAR.CommandText = "DELETE FROM EGRESOS"
            queryBORRAR.ExecuteScalar()
            queryBORRAR.CommandText = "DELETE FROM DetalleFira"
            queryBORRAR.ExecuteScalar()

            total = f1.Length
            For I As Integer = 0 To f1.Length - 1
                If f1(I).Name = "XCopy397609052017.csv.pol" Then
                    I = I
                End If
                porcentaje = (I + 1) / total
                insertaDatos(f1(I).FullName, f1(I).Name)
                Console.WriteLine(porcentaje.ToString("p") & " " & f1(I).Name)
            Next
            queryBORRAR.CommandText = "DELETE FROM EGRESOS where importeegreso = 0"
            queryBORRAR.ExecuteScalar()

            Console.WriteLine("Terminado")
        Else
            Console.WriteLine("Terminado Con Errores")
        End If
    End Sub

    Private Sub RevisaArchivos()
        Dim mes As Integer
        Dim Archivo As String = ""
        Dim ArchivoII As String = ""
        Dim dirOrg As New DirectoryInfo(RutaFira)
        Dim di1 As New DirectoryInfo(RutaFiraMes)
        Dim f1 As FileInfo() = dirOrg.GetFiles()
        ''COPIA LOS DATOS DEL MES
        For I As Integer = 0 To f1.Length - 1
            If f1(I).Name = "397609052017.csv" Then
                I = I
            End If
            Archivo = f1(I).Name
            If Mid(f1(I).Name, 1, 1) = "3" Or Mid(f1(I).Name, 1, 1) = "2" Or Mid(f1(I).Name, 1, 1) = "1" Then
                mes = Val(Mid(f1(I).Name, 7, 2))
                If Not Directory.Exists(RutaFira & "\" & MonthName(mes)) Then
                    Directory.CreateDirectory(RutaFira & "\" & MonthName(mes))
                End If
                f1(I).CopyTo(RutaFira & "\" & MonthName(mes) & "\" & Archivo, True)
                f1(I).Delete()
            ElseIf Mid(f1(I).Name, 1, 7) = "S976SDO" Then
                mes = Val(Mid(f1(I).Name, 13, 2))
                If Not Directory.Exists(RutaFira & "\" & MonthName(mes) & "\EdoCta") Then
                    Directory.CreateDirectory(RutaFira & "\" & MonthName(mes) & "\EdoCta")
                End If
                f1(I).CopyTo(RutaFira & "\" & MonthName(mes) & "\EdoCta\" & Archivo, True)
                f1(I).Delete()
            End If
        Next
        ''COPIA LOS DATOS DEL MES
        Dim Cobros As New System.IO.StreamWriter(RutaPol & "\CXSG.csv", System.IO.FileMode.Create, Text.Encoding.GetEncoding(1252))
        f1 = di1.GetFiles("*.csv")
        total = f1.Length
        For I As Integer = 0 To f1.Length - 1
            If f1(I).Name = "397609052017.csv" Then
                I = I
            End If
            porcentaje = (I + 1) / total
            Archivo = f1(I).Name
            If Mid(f1(I).Name, 1, 1) = "3" Or Mid(f1(I).Name, 1, 1) = "2" Or Mid(f1(I).Name, 1, 1) = "1" Then
                Archivo = "Copy" & f1(I).Name
                ArchivoII = "XCopy" & f1(I).Name
                f1(I).CopyTo(RutaPol & "\" & Archivo)
                Call Formatea(Archivo, RutaPol & "\" & Archivo, Cobros)
                File.Delete(RutaPol & "\" & Archivo)
                If Mid(f1(I).Name, 1, 1) = "2" Or Mid(f1(I).Name, 1, 1) = "1" Then
                    Call SacaDatos(ArchivoII, RutaPol & "\" & ArchivoII)
                Else
                    Call SacaDatos(ArchivoII, RutaPol & "\" & ArchivoII)
                End If

                File.Delete(RutaPol & "\" & ArchivoII)
            Else
                File.Delete(f1(I).FullName)
            End If
            Console.WriteLine(porcentaje.ToString("p") & " " & f1(I).Name)
        Next
        Cobros.Close()
    End Sub

    Private Sub EnviaConfirmacion(ByVal Para As String, ByVal Mensaje As String, ByVal Asunto As String)
        'Dim Mensage As New MailMessage("Interno@Finagil.com.mx", Trim(Para), Trim(Asunto), Mensaje)
        ''Dim Cliente As New SmtpClient("192.168.10.23")
        'Dim Cliente As New SmtpClient("smtp01.cmoderna.com", 26)
        'Try
        '    Cliente.Send(Mensage)
        '    'Enviado = True
        'Catch ex As Exception
        '    'Enviado = False
        '    ReportError(ex)
        'End Try
    End Sub

    Private Sub ReportError(ByVal ex As Exception)
        ErrorControl.WriteEntry(ex.Message, EventLogEntryType.Error)
    End Sub

    Private Sub SacaDatos(ByVal Archivo As String, ByVal Completo As String)
        Dim Aux As String = ""
        Dim Importe As String = ""
        Dim ImporteT As Double = 0
        Dim CargoAbono As String = ""
        Dim Concepto As String = ""
        Dim ConceptoX As String = ""
        Dim Linea As String
        Dim fechaPOL As Date = Now
        Dim LineaX As String()
        Dim Priemera As Integer = 0
        Dim i As Integer = 0
        Dim Sfile As New StreamReader(Completo, Text.Encoding.GetEncoding(1252))
        Dim Nuevo As New System.IO.StreamWriter(RutaPol & "\" & Archivo & ".pol", System.IO.FileMode.Create, Text.Encoding.GetEncoding(1252))

        'Try
        Dim TipoArch As String = Mid(Archivo, 6, 1)
        While Not Sfile.EndOfStream
            Priemera += 1
            Linea = Sfile.ReadLine()
            If Priemera = 1 Then
                'i = InStr(Linea, "CLAVE BANCO")
                'If i <= 0 Then Exit While
                Linea = Sfile.ReadLine()
            End If
            LineaX = (Linea.Split("|"))
            If LineaX(20) = "1492161" Then
                LineaX(20) = "1492161"
            End If
            Importe = "XXXXXXX"
            CargoAbono = "XXXXXX"
            Concepto = "XXXXXXX"
            If Trim(LineaX(42)) <> "" Then CargoAbono = 0
            If Trim(LineaX(43)) <> "" Then CargoAbono = 1
            If Trim(LineaX(42)) <> "" Then Importe = LineaX(49)
            If Trim(LineaX(43)) <> "" Then Importe = Trim(LineaX(50))
            If Trim(LineaX(45)) <> "" Then Concepto = Trim(LineaX(45))
            If Trim(LineaX(45)) = "" Then Concepto = Trim(LineaX(46))
            If Trim(LineaX(28)) <> "" Then fechaPOL = Trim(LineaX(29))

            If Trim(LineaX(42)) = "" And Trim(LineaX(43)) = "" And Trim(LineaX(45)) = "501" Then CargoAbono = 0
            If Trim(LineaX(42)) = "" And Trim(LineaX(43)) = "" And Trim(LineaX(45)) = "69" Then CargoAbono = 1
            If Trim(LineaX(42)) = "" And Trim(LineaX(43)) = "" And Trim(LineaX(45)) = "501" Then Importe = Trim(LineaX(49))
            If Trim(LineaX(42)) = "" And Trim(LineaX(43)) = "" And Trim(LineaX(45)) = "69" Then Importe = Trim(LineaX(50))

            If Trim(LineaX(42)) = "" And Trim(LineaX(43)) = "" And Trim(LineaX(49)) <> "" Then CargoAbono = 0
            If Trim(LineaX(42)) = "" And Trim(LineaX(43)) = "" And Trim(LineaX(50)) <> "" Then CargoAbono = 1
            If Trim(LineaX(42)) = "" And Trim(LineaX(43)) = "" And Trim(LineaX(49)) <> "" Then Importe = Trim(LineaX(49))
            If Trim(LineaX(42)) = "" And Trim(LineaX(43)) = "" And Trim(LineaX(50)) <> "" Then Importe = Trim(LineaX(50))

            Select Case Concepto
                Case 10, 39 ' CAPITAL
                    If sacatipoCredito(Trim(LineaX(20))) = "R" Then
                        Concepto = 76
                    Else
                        Concepto = 68
                    End If
                Case 11, 19, 906, 13, 14, 501 'PAGO INTERESES
                    Concepto = 69
                Case 12 'PAGO FINANCIADOS
                    Concepto = 70
                Case 16, 17 'SE EXCLUYE garantia
                    Concepto = "XX"
                Case 69 'FINANCIAMIENTO ADICIONAL
                    Concepto = "699"
                Case 8 'Fondeo de fira
                Case 18 'ejercer garantia fega
                Case 693 'PAGO DE GARANTIA
                Case 904 'PAGO DE GARANTIA
                Case 903 'penalizacion
                Case 221 'APOYO EN TASA INTERES Z25
                Case 1 '1	MINISTRACION POR TRATAMIENTO
                Case 170 'APOYO EN TASA INTERES Z25
                Case 5030 ' APOYO INFORMATIVO
                    Concepto = "XX"
                Case 5030 ' APOYO INFORMATIVO
                    Concepto = "XX"
                Case Else
                    Errores = True
                    Dim ERRR As New System.IO.StreamWriter(RutaFira & "\Errores.txt", System.IO.FileMode.Append, Text.Encoding.GetEncoding(1252))
                    ERRR.WriteLine("Nuevo Momimiento " & Concepto)
                    Console.WriteLine("Nuevo Momimiento " & Concepto)
                    EnviaConfirmacion("ecacerest@lamoderna.com.mx", "Nuevo Movimiento" & Concepto, "Nuevo Momimiento " & Concepto & " " & Archivo)
                    Concepto = "XX"
                    ERRR.Close()
                    ERRR.Dispose()

            End Select
            If Trim(LineaX(47)).ToUpper = "GARANTÍA SIN FONDEO" Or Trim(LineaX(47)).ToUpper = "GARANTIA SIN FONDEO" Then
                Concepto = "XX"
            End If
            If Concepto <> "XX" Then
                ConceptoX = "CREDITO FIRA ASOCIADO " & Trim(LineaX(20))
                If Concepto = 699 Then
                    Concepto = "70"
                    ImporteT = ImporteT - Val(Importe)
                    Nuevo.WriteLine(Concepto & ",1," & Trim(Importe) & sacaDatosExtra(Trim(LineaX(20)), Archivo) & ",'" & fechaPOL.ToString("yyyyMMdd") & "',11,'" & ConceptoX & "'")
                Else
                    ImporteT = ImporteT + Val(Importe)
                    If TipoArch = "3" Then
                        Nuevo.WriteLine(Concepto & ",0," & Trim(Importe) & sacaDatosExtra(Trim(LineaX(20)), Archivo) & ",'" & fechaPOL.ToString("yyyyMMdd") & "',11,'" & ConceptoX & "'")
                    Else
                        If ValidaIDcredito(Trim(LineaX(20)), Archivo, LineaX) = True Then
                            Nuevo.WriteLine(Trim(LineaX(20)) & ",1,'" & fechaPOL.ToString("yyyyMMdd") & "','" & fechaPOL.ToString("yyyyMMdd") & "',0,0,0," & Trim(Importe) & "," & Trim(Importe) & ",0,0,0,0,0")
                        End If
                    End If
                End If
            End If
        End While
        'If ImporteT = 0 Then NO SE NECESITA ASIENTO DE BANCOS
        '    ' no se escribe nada
        'ElseIf ImporteT > 0 Then
        '    Nuevo.WriteLine("99,1," & ImporteT & ",'','','','" & fechaPOL.ToString("yyyyMMdd") & "',11,''")
        'Else
        '    Nuevo.WriteLine("99,0," & Math.Abs(ImporteT) & ",'','','','" & fechaPOL.ToString("yyyyMMdd") & "',11,''")
        'End If
        ImporteT = 0

        'Catch ex As Exception
        '   Console.Write(ex.ToString)
        'End Try
        Sfile.Close()
        Nuevo.Close()
    End Sub

    Private Sub Formatea(ByVal Archivo As String, ByVal Completo As String, ByRef Cobros As StreamWriter)
        Dim Aux As String = ""
        Dim letra As String = ""
        Dim Linea As String
        Dim LineaX() As String
        Dim Bandera As Boolean = False
        Dim Sfile As New StreamReader(Completo, Text.Encoding.GetEncoding(1252))
        Dim Nuevo As New System.IO.StreamWriter(RutaPol & "\X" & Archivo, System.IO.FileMode.Create, Text.Encoding.GetEncoding(1252))
        Dim Xx As Integer = 0
        'Try
        While Not Sfile.EndOfStream
            Linea = Sfile.ReadLine()
            LineaX = Linea.Split(",")
            If Xx = 0 Then
                Cobros.WriteLine(Linea)
                Xx += 1
            End If
            If LineaX(46) = "16" Then
                Cobros.WriteLine(Linea)
            End If

            Aux = ""
            For x As Integer = 1 To Linea.Length
                letra = Mid(Linea, x, 1)
                If letra = """" Then
                    If Bandera = False Then
                        Bandera = True
                    Else
                        Bandera = False
                    End If
                End If
                If letra = "," And Bandera = False Then letra = "|"
                Aux = Aux & letra
            Next
            Nuevo.WriteLine(Aux)

        End While
        'Catch ex As Exception
        '    Console.Write(ex.ToString)
        '    EnviaConfirmacion("ecacerest@lamoderna.com.mx", ex.Message, "Error " & Now)
        'End Try
        Sfile.Close()
        Nuevo.Close()
    End Sub

    Function sacaDatosExtra(ByVal id As Double, ByVal archivo As String) As String
        Dim Nuevo As New System.IO.StreamWriter(RutaFira & "\Errores.txt", System.IO.FileMode.Append, Text.Encoding.GetEncoding(1252))

        Dim cad As String = ""
        Dim cli As String = ""
        Dim query As New SqlClient.SqlCommand
        Dim Datos As SqlClient.SqlDataReader
        If Conn.State = ConnectionState.Closed Then Conn.Open()
        query.Connection = Conn
        query.CommandType = CommandType.Text
        query.CommandText = "SELECT     dbo.PasivoFIRA.Anexo, dbo.Sucursales.Segmento_Negocio, TipoCredito, clientes.cliente" _
                            & " FROM         dbo.PasivoFIRA INNER JOIN" _
                            & " dbo.Clientes ON dbo.PasivoFIRA.Cliente = dbo.Clientes.Cliente INNER JOIN" _
                            & " dbo.Sucursales ON dbo.Clientes.Sucursal = dbo.Sucursales.ID_Sucursal" _
                            & " where IDCredito = " & id
        Datos = query.ExecuteReader
        If Datos.HasRows Then
            Datos.Read()
            cad = ",'" & Trim(Datos.Item(0)) & "','" & Trim(Datos.Item(2) & "','" & Trim(Datos.Item(3))) & "'"
        Else
            If EsRefaccionario(id, archivo) = False Then
                If EsGarantia(id, archivo) = False Then
                    cad = ",'*********************************'"
                    Console.WriteLine("Sin Datos EXTRA " & id & " " & archivo & " " & Now)
                    Nuevo.WriteLine("Sin Datos EXTR A" & id & " " & archivo & " " & Now)
                    Errores = True
                    EnviaConfirmacion("ecacerest@lamoderna.com.mx", "Sin Datos EXTRA" & id & " " & archivo & " " & Now, "Sin Datos " & id & " " & archivo & " " & Now)
                    EnviaConfirmacion("cjuarezr@finagil.com.mx", "Sin Datos EXTRA" & id & " " & archivo & " " & Now, "Sin Datos " & id & " " & archivo & " " & Now)
                End If
            End If
        End If
        Datos.Close()
        Nuevo.Close()
        Return cad
    End Function

    Sub insertaDatos(ByVal arch As String, ByVal nom As String)
        Dim Nuevo As New System.IO.StreamReader(arch, Text.Encoding.GetEncoding(1252))
        Dim query As New SqlClient.SqlCommand
        Dim Linea As String
        Dim id As Double = 0

        If Conn.State = ConnectionState.Closed Then Conn.Open()

        'Try
        If Mid(nom, 6, 1) = "1" Or Mid(nom, 6, 1) = "2" Then
            While Not Nuevo.EndOfStream
                Try
                    Linea = Nuevo.ReadLine
                    id = Trim(Left(Linea, 7))
                    If InStr(Linea, "XXXXXXX") > 0 Then
                        Continue While
                    End If
                    If EsRefaccionario(id) = False And EsGarantia(id) = False Then
                        query.Connection = Conn
                        query.CommandType = CommandType.Text
                        query.CommandText = "insert into detalleFira  VALUES(" & Linea & ")"
                        query.ExecuteScalar()
                    End If
                Catch ex As Exception
                    Console.Write(ex.ToString)
                    EnviaConfirmacion("ecacerest@lamoderna.com.mx", ex.Message, "Error " & Now)
                End Try
            End While
        Else
            While Not Nuevo.EndOfStream
                Try
                    Linea = Nuevo.ReadLine
                    'If InStr(Linea, "XXXXXXX") > 0 Then
                    '    Continue While
                    'End If
                    id = Trim(Left(Right(Linea, 8), 7))
                    If EsRefaccionario(id) = False And EsGarantia(id) = False And Mid(Linea, 1, 4) <> "904," And Mid(Linea, 1, 4) <> "170," And Mid(Linea, 1, 4) <> "221," Then
                        query.Connection = Conn
                        query.CommandType = CommandType.Text
                        query.CommandText = "insert into egresos (CLAVEEGRESO, CARGOABONO, IMPORTEEGRESO, ANEXO,TIPOCREDITO,CLIENTE,FECHAEGRESO,BANCO,CONCEPTO) VALUES(" & Linea & ")"
                        query.ExecuteScalar()
                    End If
                Catch ex As Exception
                    Console.Write(ex.ToString)
                    EnviaConfirmacion("ecacerest@lamoderna.com.mx", ex.Message, "Error " & Now)
                End Try
            End While
        End If


        Nuevo.Close()


    End Sub

    Function sacatipoCredito(ByVal id As Double) As String
        Dim Nuevo As New System.IO.StreamWriter(RutaFira & "\Errores.txt", System.IO.FileMode.Append, Text.Encoding.GetEncoding(1252))

        Dim cad As String = ""
        Dim query As New SqlClient.SqlCommand
        Dim Datos As SqlClient.SqlDataReader
        If Conn.State = ConnectionState.Closed Then Conn.Open()
        query.Connection = Conn
        query.CommandType = CommandType.Text
        query.CommandText = "SELECT     TipoCredito" _
                            & " FROM         dbo.PasivoFIRA INNER JOIN" _
                            & " dbo.Clientes ON dbo.PasivoFIRA.Cliente = dbo.Clientes.Cliente INNER JOIN" _
                            & " dbo.Sucursales ON dbo.Clientes.Sucursal = dbo.Sucursales.ID_Sucursal" _
                            & " where IDCredito = " & id
        Datos = query.ExecuteReader
        If Datos.HasRows Then
            Datos.Read()
            cad = Trim(Datos.Item(0))
        Else
            If EsRefaccionario(id) = False Then
                If EsGarantia(id) = False Then
                    cad = ""
                    Errores = True
                    Console.WriteLine("Sin Datos PasivoFIRA idcredito" & id & " " & Now)
                    Nuevo.WriteLine("Sin Datos PasivoFIRA idcredito " & id & " " & Now)
                    EnviaConfirmacion("ecacerest@lamoderna.com.mx", "Sin Datos PasivoFIRA idcredito" & id & " " & Now, "Sin Datos PasivoFIRA" & id & " " & Now)
                    EnviaConfirmacion("cjuarezr@finagil.com.mx", "Sin Datos PasivoFIRA idcredito" & id & " " & Now, "Sin Datos PasivoFIRA" & id & " " & Now)
                End If
            End If
        End If
        Datos.Close()
        Nuevo.Close()
        Return cad
    End Function

    Function ValidaIDcredito(ByVal id As Double, ByVal archivo As String, Optional ByVal LineaX As String() = Nothing) As Boolean
        Dim Nuevo As New System.IO.StreamWriter(RutaFira & "\Errores.txt", System.IO.FileMode.Append, Text.Encoding.GetEncoding(1252))

        Dim cad As Boolean = True
        Dim cli As String = ""
        Dim query As New SqlClient.SqlCommand
        Dim Datos As SqlClient.SqlDataReader
        If Conn.State = ConnectionState.Closed Then Conn.Open()
        query.Connection = Conn
        query.CommandType = CommandType.Text
        query.CommandText = "SELECT     dbo.PasivoFIRA.Anexo" _
                            & " FROM         dbo.PasivoFIRA " _
                            & " where IDCredito = " & id
        Datos = query.ExecuteReader
        If Datos.HasRows Then
            Datos.Read()
            cad = True
        Else
            If EsRefaccionario(id, archivo, LineaX) = False Then
                If EsGarantia(id, archivo, LineaX) = False Then
                    cad = False
                    Console.WriteLine("Sin Datos idcredito" & id & " " & archivo & " " & Now)
                    Nuevo.WriteLine("Sin Datos idcredito" & id & " " & archivo & " " & Now)
                    EnviaConfirmacion("ecacerest@lamoderna.com.mx", "Sin Datos idcredito" & id & " " & archivo & " " & Now, "Sin Datos " & id & " " & archivo & " " & Now)
                    EnviaConfirmacion("cjuarezr@finagil.com.mx", "Sin Datos idcredito" & id & " " & archivo & " " & Now, "Sin Datos " & id & " " & archivo & " " & Now)
                    Errores = True
                End If
            End If
        End If
        Datos.Close()
        Nuevo.Close()
        Return cad
    End Function

    Sub InsertaDatosPasivoFira()
        Try
            Dim query As New SqlClient.SqlCommand
            Dim queryInsert As New SqlClient.SqlCommand
            Dim Datos As SqlClient.SqlDataReader
            If Conn.State = ConnectionState.Closed Then Conn.Open()
            If Conn2.State = ConnectionState.Closed Then Conn2.Open()
            query.Connection = Conn
            queryInsert.Connection = Conn2
            query.CommandType = CommandType.Text
            queryInsert.CommandType = CommandType.Text
            query.CommandText = "SELECT * from Vw_IdCredito where idCreditoPasivo = ''"
            Datos = query.ExecuteReader
            While Datos.Read
                queryInsert.CommandText = "INSERT INTO PasivoFIRA "
                If Datos.Item("Tipo") = "Avio" Then
                    queryInsert.CommandText = queryInsert.CommandText & "SELECT IDCredito, 'A', Cliente, Anexo, Ciclo, Tipta, Tasas, DiferencialFINAGIL, 0.08, 0,0,'SIMFAA','','S',0,0,0,0 FROM Avios WHERE IDCredito*1 = '" & Datos.Item("IDcREDITO") & "' AND IDCredito NOT IN (SELECT IDCredito FROM PasivoFIRA) ORDER BY IDCredito*1"
                Else
                    queryInsert.CommandText = queryInsert.CommandText & "SELECT IDCredito, 'R', Cliente, Anexo, '', Tipta, Tasas, Difer, 0.08, 0,0,'SIMFAA','','S',0,0,0,0 FROM Anexos WHERE IDCredito = '" & Datos.Item("IDcREDITO") & "' AND IDCredito NOT IN (SELECT IDCredito FROM PasivoFIRA) ORDER BY IDCredito*1"
                End If
                Try
                    queryInsert.ExecuteScalar()
                Catch ex As Exception
                    Console.Write(ex.ToString)
                    EnviaConfirmacion("ecacerest@lamoderna.com.mx", ex.Message, "Error " & Now)
                End Try

            End While
        Catch ex As Exception
            Console.Write(ex.ToString)
            EnviaConfirmacion("ecacerest@lamoderna.com.mx", ex.Message, "Error " & Now)
        Finally
            Conn.Close()
            Conn2.Close()
        End Try


    End Sub

    Sub MantieneCarpetas()

        If Not Directory.Exists(RutaFira) Then
            Directory.CreateDirectory(RutaFira)
        End If
        If Not Directory.Exists(RutaPol) Then
            Directory.CreateDirectory(RutaPol)
        Else
            Directory.Delete(RutaPol, True)
            Directory.CreateDirectory(RutaPol)
        End If
        If Not Directory.Exists(RutaFira & "\" & MonthName(fecha.Month)) Then
            Directory.CreateDirectory(RutaFira & "\" & MonthName(fecha.Month))
        End If
    End Sub

    Function EsRefaccionario(ByVal id As Double, Optional ByVal archivo As String = "", Optional ByVal LineaX As String() = Nothing) As Boolean
        Dim mensaje As String = ""
        Dim cad As Boolean = True
        Dim query As New SqlClient.SqlCommand
        Dim queryII As New SqlClient.SqlCommand
        Dim Datos As SqlClient.SqlDataReader
        If Conn3.State = ConnectionState.Closed Then Conn3.Open()
        query.Connection = Conn3
        query.CommandType = CommandType.Text
        query.CommandText = "SELECT * " _
                            & " FROM FIRArefaccionarios" _
                            & " where IDCredito = " & id
        Datos = query.ExecuteReader
        If Datos.HasRows Then
            Datos.Read()
            If Datos.Item("enviado") = False And Not LineaX Is Nothing Then
                For x As Integer = 0 To LineaX.Length - 1
                    mensaje = mensaje & "|" & LineaX(x)
                Next
                EnviaConfirmacion("ecacerest@lamoderna.com.mx", mensaje, "Credito Refaccionario" & id & " " & Now)
                EnviaConfirmacion("vcruz@finagil.com.mx", mensaje, "Credito Refaccionario" & id & " " & Now)
                EnviaConfirmacion("cjuarezr@finagil.com.mx", mensaje, "Credito Refaccionario" & id & " " & Now)
                EnviaConfirmacion("crivera@finagil.com.mx", mensaje, "Credito Refaccionario" & id & " " & Now)
                If Conn4.State = ConnectionState.Closed Then Conn4.Open()
                queryII.Connection = Conn4
                queryII.CommandText = "update FIRArefaccionarios set enviado = 1 where idcredito = " & id
                queryII.ExecuteNonQuery()
            End If
            cad = True
        Else
            cad = False
        End If
        Datos.Close()
        Return cad
    End Function

    Function EsGarantia(ByVal id As Double, Optional ByVal archivo As String = "", Optional ByVal LineaX As String() = Nothing) As Boolean
        Dim mensaje As String = ""
        Dim cad As Boolean = True
        Dim query As New SqlClient.SqlCommand
        Dim queryII As New SqlClient.SqlCommand
        Dim Datos As SqlClient.SqlDataReader
        If Conn3.State = ConnectionState.Closed Then Conn3.Open()
        query.Connection = Conn3
        query.CommandType = CommandType.Text
        query.CommandText = "select idgarantia from FIRArefaccionarios where idgarantia = " & id _
                            & " union " _
                            & "select idgarantia from Avios  where idgarantia = " & id
        Datos = query.ExecuteReader
        If Datos.HasRows Then
            Datos.Read()
            cad = True
        Else
            cad = False
        End If
        Datos.Close()
        Return cad
    End Function

    Sub SaldosInsolutos()
        ' Este programa debe tomar todos los contratos activos con fecha de contratación menor o igual a la
        ' fecha de proceso.   También debe tomar la tabla de amortización de todos los contratos obtenidos 
        ' con el criterio anterior, tanto la del equipo como la del seguro.   Aunque esto creará un dataset con
        ' muchos registros, por otra parte permitirá mantener abierta la conexión únicamente durante el
        ' tiempo que tarde en traer dicha información de la base de datos.

        ' Adicionalmente deberá traer todas las facturas no pagadas de los contratos activos con fecha de
        ' contratación menor o igual a la fecha de proceso.   Esto con el fin de determinar y excluir aquellos
        ' contratos con rentas vencidas a más de 89 días.

        ' Declaración de variables de conexión ADO .NET

        Dim cnAgil As New SqlConnection("Server=SERVER-RAID2; DataBase=production; User ID=User_PRO; pwd=User_PRO2015")
        Dim TA As New ProductionDataSetTableAdapters.ZSaldosInsolutosTableAdapter
        Dim cm1 As New SqlCommand()
        Dim cm2 As New SqlCommand()
        Dim cm3 As New SqlCommand()
        Dim cm4 As New SqlCommand()
        Dim cm5 As New SqlCommand()
        Dim dsAgil As New DataSet()
        Dim dsReporte As New DataSet()
        Dim daAnexos As New SqlDataAdapter(cm1)
        Dim daEdoctav As New SqlDataAdapter(cm2)
        Dim daEdoctas As New SqlDataAdapter(cm3)
        Dim daEdoctao As New SqlDataAdapter(cm4)
        Dim daFacturas As New SqlDataAdapter(cm5)
        ''¿Dim dtReporte As New DataTable("Reporte")
        Dim dtReporte As New ProductionDataSet.ZSaldosInsolutosDataTable
        'Dim dtFinal As New DataTable()
        Dim drEdoctav As DataRow()
        Dim drEdoctas As DataRow()
        Dim drEdoctao As DataRow()
        Dim drFacturas As DataRow()
        ''Dim drRegistros As DataRow()
        Dim drAnexo As DataRow
        ''Dim drRegistro As DataRow
        Dim drReporte As DataRow
        Dim relAnexoEdoctav As DataRelation
        Dim relAnexoEdoctas As DataRelation
        Dim relAnexoEdoctao As DataRelation
        Dim relAnexoFacturas As DataRelation

        ' Declaración de variables de datos

        Dim cAnexo As String
        Dim cCusnam As String
        Dim cFecha As String
        Dim cFechacon As String
        Dim cFondeo As String
        Dim cForca As String
        Dim cFvenc As String
        Dim cTipar As String
        Dim cTipo As String
        Dim cTipta As String
        Dim nAmorin As Decimal
        Dim nBonifica As Decimal
        Dim nCarteraEquipo As Decimal = 0
        Dim nCarteraSeguro As Decimal = 0
        Dim nCarteraOtros As Decimal = 0
        Dim nCounter As Integer
        Dim nMaxCounter As Integer = 100
        Dim nDifer As Decimal
        Dim nImpDG As Decimal = 0
        Dim nImpEq As Decimal
        Dim nImpRD As Decimal
        Dim nInteresEquipo As Decimal = 0
        Dim nInteresSeguro As Decimal = 0
        Dim nInteresOtros As Decimal = 0
        Dim nIvaDG As Decimal = 0
        Dim nIvaDiferido As Decimal
        Dim nIvaEq As Decimal
        Dim nIvaRD As Decimal
        Dim nPlazo As Byte
        Dim nPorieq As Decimal
        Dim nRtasD As Decimal
        Dim nSaldoEquipo As Decimal = 0
        Dim nSaldoSeguro As Decimal = 0
        Dim nSaldoOtros As Decimal = 0
        Dim nTasas As Decimal

        ' Declaración de variables de Crystal Reports

        ''Dim newrptRepsaldo As New rptRepsaldo()
        ''btnProcesar.Enabled = False
        ''cFecha = DTOC(dtpProcesar.Value)
        cFecha = Date.Now.ToString("yyyyMMdd")

        ' Primero creo la tabla Reporte que será la base del reporte

        ''dtReporte.Columns.Add("Tipo", Type.GetType("System.String"))
        ''dtReporte.Columns.Add("Contrato", Type.GetType("System.String"))
        ''dtReporte.Columns.Add("Cliente", Type.GetType("System.String"))
        ''dtReporte.Columns.Add("Termina", Type.GetType("System.String"))
        ''dtReporte.Columns.Add("Mpt", Type.GetType("System.String"))
        ''dtReporte.Columns.Add("SaldoEquipo", Type.GetType("System.Decimal"))
        ''dtReporte.Columns.Add("IvaDiferido", Type.GetType("System.Decimal"))
        ''dtReporte.Columns.Add("Deposito", Type.GetType("System.Decimal"))
        ''dtReporte.Columns.Add("SaldoSeguro", Type.GetType("System.Decimal"))
        ''dtReporte.Columns.Add("SaldoOtros", Type.GetType("System.Decimal"))
        ''dtReporte.Columns.Add("SaldoTotal", Type.GetType("System.Decimal"))
        ''dtReporte.Columns.Add("InteresEquipo", Type.GetType("System.Decimal"))
        ''dtReporte.Columns.Add("CarteraEquipo", Type.GetType("System.Decimal"))

        ' Este Stored Procedure trae todos los contratos activos con fecha de contratación menor o igual
        ' a la de proceso

        With cm1
            .CommandType = CommandType.StoredProcedure
            .CommandText = "GeneProv1"
            .Connection = cnAgil
            .Parameters.Add("@FechaFin", SqlDbType.NVarChar)
            .Parameters(0).Value = cFecha
        End With

        ' Este Stored Procedure trae la tabla de amortización del equipo de todos los contratos activos
        ' con fecha de contratación menor o igual a la de proceso

        With cm2
            .CommandType = CommandType.StoredProcedure
            .CommandText = "GeneProv2"
            .Connection = cnAgil
            .Parameters.Add("@FechaFin", SqlDbType.NVarChar)
            .Parameters(0).Value = cFecha
        End With

        ' Este Stored Procedure trae la tabla de amortización del seguro de todos los contratos activos
        ' con fecha de contratación menor o igual a la de proceso

        With cm3
            .CommandType = CommandType.StoredProcedure
            .CommandText = "GeneProv3"
            .Connection = cnAgil
            .Parameters.Add("@FechaFin", SqlDbType.NVarChar)
            .Parameters(0).Value = cFecha
        End With

        ' Este Stored Procedure trae la tabla de amortización de otros adeudos de todos los contratos activos
        ' con fecha de contratación menor o igual a la de proceso

        With cm4
            .CommandType = CommandType.StoredProcedure
            .CommandText = "GeneProv4"
            .Connection = cnAgil
            .Parameters.Add("@FechaFin", SqlDbType.NVarChar)
            .Parameters(0).Value = cFecha
        End With

        ' Este Stored Procedure trae todas las facturas no pagadas de todos los contratos activos con fecha de
        ' contratación menor o igual a la de proceso

        With cm5
            .CommandType = CommandType.StoredProcedure
            .CommandText = "CalcAnti1"
            .Connection = cnAgil
            .Parameters.Add("@Fecha", SqlDbType.NVarChar)
            .Parameters(0).Value = cFecha
        End With

        ' Llenar el DataSet a través del DataAdapter, lo cual abre y cierra la conexión

        daAnexos.Fill(dsAgil, "Anexos")
        daEdoctav.Fill(dsAgil, "Edoctav")
        daEdoctas.Fill(dsAgil, "Edoctas")
        daEdoctao.Fill(dsAgil, "Edoctao")
        daFacturas.Fill(dsAgil, "Facturas")

        ' Establecer la relación entre Anexos y Edoctav

        relAnexoEdoctav = New DataRelation("AnexoEdoctav", dsAgil.Tables("Anexos").Columns("Anexo"), dsAgil.Tables("Edoctav").Columns("Anexo"))
        dsAgil.EnforceConstraints = False
        dsAgil.Relations.Add(relAnexoEdoctav)

        ' Establecer la relación entre Anexos y Edoctas

        relAnexoEdoctas = New DataRelation("AnexoEdoctas", dsAgil.Tables("Anexos").Columns("Anexo"), dsAgil.Tables("Edoctas").Columns("Anexo"))
        dsAgil.EnforceConstraints = False
        dsAgil.Relations.Add(relAnexoEdoctas)

        ' Establecer la relación entre Anexos y Edoctao

        relAnexoEdoctao = New DataRelation("AnexoEdoctao", dsAgil.Tables("Anexos").Columns("Anexo"), dsAgil.Tables("Edoctao").Columns("Anexo"))
        dsAgil.EnforceConstraints = False
        dsAgil.Relations.Add(relAnexoEdoctao)

        ' Establecer la relación entre Anexos y Facturas

        relAnexoFacturas = New DataRelation("AnexoFacturas", dsAgil.Tables("Anexos").Columns("Anexo"), dsAgil.Tables("Facturas").Columns("Anexo"))
        dsAgil.EnforceConstraints = False
        dsAgil.Relations.Add(relAnexoFacturas)

        For Each drAnexo In dsAgil.Tables("Anexos").Rows

            cAnexo = drAnexo("Anexo")
            cCusnam = drAnexo("Descr")
            cTipar = drAnexo("Tipar")
            cTipo = drAnexo("Tipar")
            nPlazo = drAnexo("Plazo")
            cFechacon = drAnexo("Fechacon")
            cFvenc = drAnexo("Fvenc")
            cFondeo = drAnexo("Fondeo")
            cTipta = drAnexo("Tipta")
            nTasas = drAnexo("Tasas")
            nDifer = drAnexo("Difer")
            cForca = drAnexo("Forca")
            nPorieq = drAnexo("Porieq")
            nImpEq = drAnexo("ImpEq")
            nAmorin = drAnexo("Amorin")
            nIvaEq = drAnexo("IvaEq")
            nRtasD = drAnexo("RtasD")
            nImpRD = drAnexo("ImpRD")
            nIvaRD = drAnexo("IvaRD")
            nImpDG = drAnexo("ImpDG")
            nIvaDG = drAnexo("IvaDG")

            drFacturas = drAnexo.GetChildRows("AnexoFacturas")

            CalcAnti(cAnexo, cFecha, nMaxCounter, nCounter, drFacturas)

            If nCounter > nMaxCounter Then
                cTipo = "V"
            End If

            ' Esta instrucción trae la tabla de amortización del Equipo única y exclusivamente del contrato
            ' que está siendo procesado

            nSaldoEquipo = 0
            nInteresEquipo = 0
            nCarteraEquipo = 0

            drEdoctav = drAnexo.GetChildRows("AnexoEdoctav")
            TraeSald(drEdoctav, cFecha, nSaldoEquipo, nInteresEquipo, nCarteraEquipo)

            ' Esta instrucción trae la tabla de amortización del Seguro única y exclusivamente del contrato
            ' que está siendo procesado

            nSaldoSeguro = 0
            nInteresSeguro = 0
            nCarteraSeguro = 0

            drEdoctas = drAnexo.GetChildRows("AnexoEdoctas")
            TraeSald(drEdoctas, cFecha, nSaldoSeguro, nInteresSeguro, nCarteraSeguro)

            ' Esta instrucción trae la tabla de amortización de otros adeudos única y exclusivamente del contrato
            ' que está siendo procesado

            nSaldoOtros = 0
            nInteresOtros = 0
            nCarteraOtros = 0

            drEdoctao = drAnexo.GetChildRows("AnexoEdoctao")
            TraeSald(drEdoctao, cFecha, nSaldoOtros, nInteresOtros, nCarteraOtros)

            ' Debe determinar si es un contrato con IVA diferido

            nIvaDiferido = 0
            If cFechacon >= "20020301" And nIvaEq > 0 Then
                nIvaDiferido = System.Math.Round(nSaldoEquipo * nPorieq / 100, 2)
            End If

            ' Debe determinar si el cliente dejó un Depósito en Garantía y netearlo contra el IVA diferido

            nBonifica = 0
            If cFechacon >= "200020901" And nImpRD > 0 And nRtasD = 0 Then
                nBonifica = System.Math.Round(nSaldoEquipo * ((nImpRD + nIvaRD) / (nImpEq - nIvaEq - nAmorin)), 2)
            End If

            If cTipar = "F" Then
                nIvaDiferido = System.Math.Round(nIvaDiferido - nBonifica, 2)
                nBonifica = 0
            End If

            ' También tengo que determinar si el cliente dejó Rentas en Depósito

            If nImpDG > 0 Then
                nBonifica = System.Math.Round(nBonifica + nImpDG + nIvaDG, 2)
            End If

            drReporte = dtReporte.NewRow()
            drReporte("Tipo") = cTipo
            drReporte("Contrato") = Mid(cAnexo, 1, 5) & "/" & Mid(cAnexo, 6, 4)
            drReporte("Cliente") = Trim(cCusnam)
            drReporte("Termina") = Mid(Termina(CTOD(cFvenc), nPlazo).ToString, 1, 10)
            drReporte("Mpt") = Mpt(Termina(CTOD(cFvenc), nPlazo), Date.Now)
            drReporte("SaldoEquipo") = nSaldoEquipo
            drReporte("IvaDiferido") = nIvaDiferido
            drReporte("Deposito") = nBonifica
            drReporte("SaldoSeguro") = nSaldoSeguro
            drReporte("SaldoOtros") = nSaldoOtros
            drReporte("SaldoTotal") = nSaldoEquipo + nIvaDiferido - nBonifica + nSaldoSeguro + nSaldoOtros
            drReporte("InteresEquipo") = nInteresEquipo
            drReporte("CarteraEquipo") = nCarteraEquipo
            dtReporte.Rows.Add(drReporte)

        Next
        TA.DeleteAll()
        TA.Update(dtReporte)
        dtReporte.AcceptChanges()

        Try

            ''dtFinal = dtReporte.Clone()
            ''drRegistros = dtReporte.Select(True, "Tipo, Contrato")

            ''For Each drRegistro In drRegistros
            ''    dtFinal.ImportRow(drRegistro)
            ''Next

            ''dsReporte.Tables.Add(dtFinal)
            ' '' Descomentar la siguiente línea en caso de que se deseara modificar el reporte rptRepSaldo
            ' '' dsReporte.WriteXml("C:\Schema4.xml", XmlWriteMode.WriteSchema)
            ''newrptRepsaldo.SetDataSource(dsReporte)
            ''newrptRepsaldo.SummaryInfo.ReportComments = Mes(cFecha)
            ''CrystalReportViewer1.ReportSource = newrptRepsaldo

        Catch eException As Exception

            ''MsgBox(eException.Message, MsgBoxStyle.Critical, "Mensaje de Error")

        End Try

        cnAgil.Dispose()
        cm1.Dispose()
        cm2.Dispose()
        cm3.Dispose()
        cm4.Dispose()
        cm5.Dispose()
    End Sub

    Public Sub CalcAnti(ByVal cAnexo As String, ByVal cFecha As String, ByRef nMaxCounter As Integer, ByRef nCounter As Integer, ByVal drFacturas As DataRow())

        ' Esta función calcula la antigüedad de todas las facturas no pagadas de este contrato, a fin de determinar
        ' la más antigüa de ellas.   Debe recibir como parámetro un DataRowCollection el cual contenga todas las
        ' facturas no pagadas de un contrato dado.

        ' Declaración de variables de conexión

        Dim drFactura As DataRow

        ' Declaración de variables de datos

        Dim cFeven As String
        Dim nDiasRetraso As Integer

        If cFecha >= "20000101" Then
            nMaxCounter = 89
        ElseIf cFecha >= "19970101" Then
            nMaxCounter = 90
        ElseIf cFecha >= "19960301" Then
            nMaxCounter = 180
        End If

        nCounter = 0

        For Each drFactura In drFacturas
            cFeven = drFactura("Feven")
            nDiasRetraso = DateDiff(DateInterval.Day, CTOD(cFeven), CTOD(cFecha))
            If nDiasRetraso > nCounter Then
                nCounter = nDiasRetraso
            End If
        Next

    End Sub

    Public Function CTOD(ByVal cFecha As String) As Date

        Dim nDia, nMes, nYear As Integer

        nDia = Val(Right(cFecha, 2))
        nMes = Val(Mid(cFecha, 5, 2))
        nYear = Val(Left(cFecha, 4))

        CTOD = DateSerial(nYear, nMes, nDia)

    End Function

    Public Function DTOC(ByVal dFecha As Date) As String

        Dim cDia, cMes, cYear, sFecha As String

        sFecha = dFecha.ToShortDateString

        cDia = Left(sFecha, 2)
        cMes = Mid(sFecha, 4, 2)
        cYear = Right(sFecha, 4)

        DTOC = cYear & cMes & cDia

    End Function

    Public Sub TraeSald(ByVal drVencimientos As DataRow(), ByVal cFeven As String, ByRef nSaldo As Decimal, ByRef nInteres As Decimal, ByRef nCartera As Decimal)

        ' Esta variable datarow contendrá los datos de 1 vencimiento a la vez, de la tabla Edoctav, Edoctas o Edoctao

        Dim drVencimiento As DataRow

        For Each drVencimiento In drVencimientos
            If (drVencimiento("Feven") >= cFeven And drVencimiento("IndRec") = "S") Or drVencimiento("Nufac") = 0 Then
                nSaldo += drVencimiento("Abcap")
                nInteres += drVencimiento("Inter")
                nCartera += drVencimiento("Abcap") + drVencimiento("Inter")
            End If
        Next
        nSaldo = System.Math.Round(nSaldo, 2)
        nInteres = System.Math.Round(nInteres, 2)
        nCartera = System.Math.Round(nCartera, 2)

    End Sub

    Public Function Termina(ByVal dInicio As Date, ByVal nPlazo As Integer) As Date

        Dim nDay As Byte
        Dim nYear As Integer
        Dim nMonth As Byte
        'Dim nAños As Integer
        'Dim nAñosb As Integer
        'Dim nLeap As Byte
        'Dim i As Integer
        'Dim nMes As Integer
        'Dim nDia As Integer

        nDay = Day(dInicio)
        nMonth = (Month(dInicio) + (nPlazo Mod 12)) - 1
        nYear = Year(dInicio) + Int(nPlazo / 12)
        If nMonth > 12 Then
            nMonth -= 12
            nYear += 1
        ElseIf nMonth = 0 Then
            nMonth = 12
            nYear -= 1
        End If
        Termina = DateSerial(nYear, nMonth, nDay)

        If nDay <> 6 And nDay <> 16 And nDay <> 20 And nDay <> 25 Then
            Termina = IIf(Month(Termina) > nMonth, DateAdd(DateInterval.Month, -1, Termina), Termina)
            Termina = DayWeek(Termina)
        End If

    End Function

    Public Function DayWeek(ByVal dFecha As Date)

        Dim nDay As Byte
        Dim nYear As Integer
        Dim nMonth As Byte
        Dim nAños As Integer
        Dim nAñosb As Integer
        Dim nLeap As Byte
        Dim i As Integer
        Dim nMes As Integer
        Dim nDia As Integer

        nDay = Day(dFecha)
        nMonth = Month(dFecha)
        nYear = Year(dFecha)

        If nMonth = 12 Then
            nMonth = 1
            nYear += 1
        Else
            nMonth += 1
        End If

        dFecha = DateSerial(nYear, nMonth, 1)
        dFecha = DateAdd(DateInterval.Day, -1, dFecha)

        nDay = Day(dFecha)
        nMonth = Month(dFecha)
        nYear = Year(dFecha)

        nAños = nYear - 1933
        nLeap = 0
        nAñosb = 0

        For i = 1933 To nYear
            nLeap = Leap(i)
            If nLeap = 1 Then
                nAñosb += 1
                nLeap = 0
            End If
        Next

        Select Case nMonth
            Case 1, 10
                nMes = 0
            Case 2, 3, 11
                nMes = 3
            Case 4, 7
                nMes = 6
            Case 5
                nMes = 1
            Case 6
                nMes = 4
            Case 8
                nMes = 2
            Case 9, 12
                nMes = 5
        End Select

        If nMonth = 2 And nDay = 29 Then
            nDay = 28
        End If
        nDia = (nAños + nAñosb + nMes + nDay) Mod 7
        If nDia = 1 Then
            dFecha = DateAdd(DateInterval.Day, -2, dFecha)
        ElseIf nDia = 0 Then
            dFecha = DateAdd(DateInterval.Day, -1, dFecha)
        End If

        DayWeek = dFecha.ToShortDateString

    End Function

    Public Function Mpt(ByVal Fecha1 As Date, ByVal Fecha2 As Date) As Integer
        Mpt = ((Year(Fecha1) * 12) + Month(Fecha1)) - ((Year(Fecha2) * 12) + Month(Fecha2))
    End Function

    Public Function Leap(ByVal nYear As Integer)
        Leap = 0
        If nYear Mod 400 = 0 Then
            Leap = 1
        ElseIf nYear Mod 100 = 0 Then
            Leap = 0
        ElseIf nYear Mod 4 = 0 Then
            Leap = 1
        End If
        Return (Leap)

    End Function

    Sub Antiguedad()

        ' Declaración de variables de conexión ADO .NET

        Dim cnAgil As New SqlConnection("Server=SERVER-RAID2; DataBase=production; User ID=User_PRO; pwd=User_PRO2015")
        Dim TA As New ProductionDataSetTableAdapters.ZantiguedadTableAdapter
        Dim cm1 As New SqlCommand()
        Dim cm2 As New SqlCommand()
        Dim cm3 As New SqlCommand()
        Dim cm4 As New SqlCommand()
        Dim cm5 As New SqlCommand()
        Dim cm6 As New SqlCommand()
        Dim dsAgil As New DataSet()
        Dim daFacturas As New SqlDataAdapter(cm1)
        Dim daPagosIniciales As New SqlDataAdapter(cm2)
        Dim daOpciones As New SqlDataAdapter(cm3)
        Dim daAnexos As New SqlDataAdapter(cm4)
        Dim daReestructurados As New SqlDataAdapter(cm5)
        Dim daSeguimientos As New SqlDataAdapter(cm6)
        Dim drFactura As DataRow
        Dim drAnexo As DataRow
        Dim drSeguimiento As DataRow
        Dim dvClientes As DataView
        Dim dvAnexos As DataView
        Dim myColArray(1) As DataColumn
        Dim dtClientes As New DataTable()
        Dim dtAnexos As New ProductionDataSet.ZantiguedadDataTable
        Dim dtReporte As New ProductionDataSet.ZantiguedadDataTable
        Dim drCliente As DataRow
        Dim drReporte As DataRow

        ' Declaración de variables de datos

        Dim cAnexo As String
        Dim cDigito As String
        Dim cCliente As String
        Dim cFecha As String
        Dim cFechaLlamada As String = ""
        Dim cFechaSeguimiento As String = ""    ' Es la fecha a partir de la cual el sistema va a contar el número de gestiones de cobranza que se hubieran realizado
        Dim cGestor As String
        Dim cInicialesPromotor As String = ""
        Dim cLetra As String
        Dim cNombre As String
        Dim cPromotor As String
        Dim cSeguimiento As String = ""
        Dim cSucursal As String
        Dim cVencida As String = ""
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim nAnexos As Integer = 0
        Dim nClientes As Integer = 0
        Dim nDiasVencido As Integer = 0
        Dim nGestion As Integer = 0
        Dim nPlazo As Byte
        Dim nSaldoFac As Decimal
        Dim nSaldoInsoluto As Decimal

        ' Declaración de variables de Crystal Reports

        ''Dim newrptRepAntig1 As rptRepAnti3
        ''Dim newrptRepAntig2 As rptRepAntig

        ''Dim cReportTitle As String
        ''Dim cReportComments As String

        cFecha = DTOC(Date.Now)

        ' Primero creo la tabla Clientes que me dará el orden del reporte, acumulando saldos por cliente

        dtClientes.Columns.Add("Cliente", Type.GetType("System.String"))
        dtClientes.Columns.Add("Col1a29", Type.GetType("System.Decimal"))
        dtClientes.Columns.Add("Col30a59", Type.GetType("System.Decimal"))
        dtClientes.Columns.Add("Col60a89", Type.GetType("System.Decimal"))
        dtClientes.Columns.Add("Col90omas", Type.GetType("System.Decimal"))
        dtClientes.Columns.Add("Promotor", Type.GetType("System.String"))
        dtClientes.Columns.Add("InicialesPromotor", Type.GetType("System.String"))
        dtClientes.Columns.Add("Gestor", Type.GetType("System.String"))
        dtClientes.Columns.Add("Sucursal", Type.GetType("System.String"))
        dtClientes.Columns.Add("Retraso", Type.GetType("System.Decimal"))           ' Guardo el máximo retraso del Cliente
        dtClientes.Columns.Add("PoN", Type.GetType("System.String"))
        myColArray(0) = dtClientes.Columns("Cliente")
        dtClientes.PrimaryKey = myColArray

        ' '' Ahora creo la tabla Anexos que será la base del reporte

        ''dtAnexos.Columns.Add("Anexo", Type.GetType("System.String"))
        ''dtAnexos.Columns.Add("Cliente", Type.GetType("System.String"))
        ''dtAnexos.Columns.Add("Nombre", Type.GetType("System.String"))
        ''dtAnexos.Columns.Add("Letra", Type.GetType("System.String"))
        ''dtAnexos.Columns.Add("Plazo", Type.GetType("System.Decimal"))
        ''dtAnexos.Columns.Add("Dia", Type.GetType("System.Decimal"))
        ''dtAnexos.Columns.Add("Saldo", Type.GetType("System.Decimal"))
        ''dtAnexos.Columns.Add("Col1a29", Type.GetType("System.Decimal"))
        ''dtAnexos.Columns.Add("Col30a59", Type.GetType("System.Decimal"))
        ''dtAnexos.Columns.Add("Col60a89", Type.GetType("System.Decimal"))
        ''dtAnexos.Columns.Add("Col90omas", Type.GetType("System.Decimal"))
        ''dtAnexos.Columns.Add("Total", Type.GetType("System.Decimal"))
        ''dtAnexos.Columns.Add("Retraso", Type.GetType("System.Decimal"))             ' Guardo el máximo retraso del Anexo
        ''dtAnexos.Columns.Add("Reestructura", Type.GetType("System.String"))
        ''dtAnexos.Columns.Add("PI", Type.GetType("System.Decimal"))
        ''dtAnexos.Columns.Add("OC", Type.GetType("System.Decimal"))
        ''dtAnexos.Columns.Add("Recursos", Type.GetType("System.String"))
        ''dtAnexos.Columns.Add("InicialesPromotor", Type.GetType("System.String"))
        ''dtAnexos.Columns.Add("Gestion", Type.GetType("System.Decimal"))
        ''dtAnexos.Columns.Add("Seguimiento", Type.GetType("System.String"))
        ''dtAnexos.Columns.Add("FechaLlamada", Type.GetType("System.String"))
        ''dtAnexos.Columns.Add("Vencida", Type.GetType("System.String"))
        ''myColArray(0) = dtAnexos.Columns("Anexo")
        ''dtAnexos.PrimaryKey = myColArray

        dtReporte = dtAnexos.Clone

        ' Este Stored Procedure trae las facturas vencidas, el pago inicial (sin el 5% Nafin) 
        ' y la opción de compra exigible de los contratos activos o terminados con saldo en rentas

        With cm1
            .CommandType = CommandType.StoredProcedure
            .CommandText = "Repantig1"
            .Connection = cnAgil
            .Parameters.Add("@Fecha", SqlDbType.NVarChar)
            .Parameters(0).Value = cFecha
        End With

        ' Este Stored Procedure regresa los contratos activos, terminados o cancelados 
        ' que ya cubrieron su pago inicial

        With cm2
            .CommandType = CommandType.StoredProcedure
            .CommandText = "PagosIni1"
            .Connection = cnAgil
        End With

        ' Este Stored Procedure regresa todas las opciones de compra exigibles y no pagadas

        With cm3
            .CommandType = CommandType.StoredProcedure
            .CommandText = "Repantig3"
            .Connection = cnAgil
        End With

        ' Este Stored Procedure regresa todas los contratos activos con fecha de contratación menor o igual
        ' a la fecha del reporte.   Esto me servirá para posteriormente verificar cuáles de estos contratos
        ' NO han cubierto sus pagos iniciales

        With cm4
            .CommandType = CommandType.StoredProcedure
            .CommandText = "Repantig4"
            .Connection = cnAgil
            .Parameters.Add("@Fecha", SqlDbType.NVarChar)
            .Parameters(0).Value = cFecha
        End With

        ' Este Stored Procedure regresa el número de los anexos activos o terminados que han tenido
        ' reestructura por capitalización de adeudo

        With cm5
            .CommandType = CommandType.StoredProcedure
            .CommandText = "Repantig5"
            .Connection = cnAgil
        End With

        ' El siguiente Command trae el seguimiento de los 30 días anteriores a la fecha del reporte y hasta la fecha del reporte

        cFechaSeguimiento = DTOC(DateAdd(DateInterval.Day, -30, CTOD(cFecha)))

        With cm6
            .CommandType = CommandType.Text
            .CommandText = "SELECT * FROM JUR_BitacoraCob WHERE Fecha >= '" & cFechaSeguimiento & "' ORDER BY Cliente, Fecha DESC, id_bitacoraCob"
            .Connection = cnAgil
        End With

        ' Llenar el DataSet a través del DataAdapter, lo cual abre y cierra la conexión

        Try
            daFacturas.Fill(dsAgil, "Facturas")
            daAnexos.Fill(dsAgil, "Anexos")
            daPagosIniciales.Fill(dsAgil, "PagosIniciales")
            daOpciones.Fill(dsAgil, "Opciones")
            daReestructurados.Fill(dsAgil, "Reestructurados")
            daSeguimientos.Fill(dsAgil, "Seguimientos")
        Catch eException As Exception
            MsgBox(eException.Message, MsgBoxStyle.Critical, "Mensaje de Error")
        End Try

        ' Tengo que definir una llave primaria para la tabla de Pagos Iniciales para que se acelere la búsqueda
        ' cuando revise si existe el pago inicial de un contrato

        myColArray(0) = dsAgil.Tables("PagosIniciales").Columns("Anexo")
        dsAgil.Tables("PagosIniciales").PrimaryKey = myColArray

        ' En una sola barrida a la tabla Facturas voy a determinar la antigüedad por Cliente y por Anexo

        For Each drFactura In dsAgil.Tables("Facturas").Rows

            cDigito = IIf(System.Math.Pow(-1, Val(Mid(drFactura("Cliente"), 5, 1))) > 0, "P", "N")
            cCliente = drFactura("Cliente")
            cNombre = drFactura("Descr")
            cAnexo = drFactura("Anexo")
            cLetra = drFactura("Letra")
            nPlazo = drFactura("Plazo")
            cPromotor = drFactura("DescPromotor")
            cInicialesPromotor = drFactura("Iniciales")
            cGestor = Trim(drFactura("Gestor"))
            cSucursal = drFactura("Sucursal")
            nSaldoInsoluto = drFactura("SaldoInsoluto")
            nSaldoFac = drFactura("SaldoFac")
            cVencida = drFactura("Vencida")

            ''If chkConsiderarDia.Checked = True Then
            nDiasVencido = DateDiff(DateInterval.Day, CTOD(drFactura("Feven")), CTOD(cFecha)) + 1
            ''Else
            ''    nDiasVencido = DateDiff(DateInterval.Day, CTOD(drFactura("Feven")), CTOD(cFecha))
            ''End If

            If nDiasVencido > 0 Then

                ' Busco el Cliente en la tabla Clientes

                drCliente = dtClientes.Rows.Find(cCliente)

                If drCliente Is Nothing Then
                    drCliente = dtClientes.NewRow()
                    drCliente("Cliente") = cCliente
                    drCliente("Col1a29") = 0
                    drCliente("Col30a59") = 0
                    drCliente("Col60a89") = 0
                    drCliente("Col90omas") = 0
                    If nDiasVencido > 89 Then
                        drCliente("Col90omas") = nSaldoFac
                    ElseIf nDiasVencido > 59 Then
                        drCliente("Col60a89") = nSaldoFac
                    ElseIf nDiasVencido > 29 Then
                        drCliente("Col30a59") = nSaldoFac
                    ElseIf nDiasVencido > 0 Then
                        drCliente("Col1a29") = nSaldoFac
                    End If
                    drCliente("Promotor") = cPromotor
                    drCliente("InicialesPromotor") = cInicialesPromotor
                    drCliente("Gestor") = cGestor
                    drCliente("Sucursal") = cSucursal
                    drCliente("Retraso") = nDiasVencido
                    drCliente("PoN") = cDigito

                    ' Cuando estoy registrando por primera vez el cliente es cuando voy a determinar el número de gestiones de cobranza que se han realizado en los últimos 30 días
                    ' así como el resultado más reciente de dichas gestiones

                    nGestion = 0
                    cSeguimiento = ""
                    cFechaLlamada = ""
                    For Each drSeguimiento In dsAgil.Tables("Seguimientos").Rows
                        If drSeguimiento("Cliente") = cCliente Then
                            nGestion = nGestion + 1
                            If nGestion = 1 Then
                                cSeguimiento = drSeguimiento("Resultado")
                                cFechaLlamada = drSeguimiento("Fecha")
                                cFechaLlamada = Mid(cFechaLlamada, 7, 2) + "/" + Mid(cFechaLlamada, 5, 2) + "/" + Mid(cFechaLlamada, 1, 4)
                            End If
                        End If
                    Next

                    dtClientes.Rows.Add(drCliente)

                Else
                    If nDiasVencido > 89 Then
                        drCliente("Col90omas") += nSaldoFac
                    ElseIf nDiasVencido > 59 Then
                        drCliente("Col60a89") += nSaldoFac
                    ElseIf nDiasVencido > 29 Then
                        drCliente("Col30a59") += nSaldoFac
                    ElseIf nDiasVencido > 0 Then
                        drCliente("Col1a29") += nSaldoFac
                    End If

                    ' La siguiente condición permite tener en el campo Retraso el máximo número de días vencidos del Cliente

                    If nDiasVencido > drCliente("Retraso") Then
                        drCliente("Retraso") = nDiasVencido
                    End If

                End If

                ' Busco el Anexo en la tabla Anexos

                drAnexo = dtAnexos.Rows.Find(cAnexo)

                If drAnexo Is Nothing Then

                    drAnexo = dtAnexos.NewRow()
                    drAnexo("Anexo") = cAnexo
                    drAnexo("Cliente") = cCliente
                    drAnexo("Nombre") = cNombre
                    drAnexo("Letra") = cLetra
                    drAnexo("Plazo") = nPlazo
                    drAnexo("Dia") = 1
                    drAnexo("Saldo") = nSaldoInsoluto
                    drAnexo("Col1a29") = 0
                    drAnexo("Col30a59") = 0
                    drAnexo("Col60a89") = 0
                    drAnexo("Col90omas") = 0
                    If nDiasVencido > 89 Then
                        drAnexo("Col90omas") = nSaldoFac
                    ElseIf nDiasVencido > 59 Then
                        drAnexo("Col60a89") = nSaldoFac
                    ElseIf nDiasVencido > 29 Then
                        drAnexo("Col30a59") = nSaldoFac
                    ElseIf nDiasVencido > 0 Then
                        drAnexo("Col1a29") = nSaldoFac
                    End If
                    drAnexo("Total") = nSaldoFac
                    drAnexo("Retraso") = nDiasVencido
                    drAnexo("Reestructura") = ""
                    drAnexo("PI") = 0
                    drAnexo("OC") = 0
                    drAnexo("Recursos") = ""
                    drAnexo("InicialesPromotor") = cInicialesPromotor
                    drAnexo("Gestion") = nGestion
                    drAnexo("Seguimiento") = Mid(Trim(cSeguimiento), 1, 50)
                    drAnexo("FechaLlamada") = cFechaLlamada
                    drAnexo("Vencida") = cVencida
                    dtAnexos.Rows.Add(drAnexo)

                Else

                    If nDiasVencido > 89 Then
                        drAnexo("Col90omas") += nSaldoFac
                    ElseIf nDiasVencido > 59 Then
                        drAnexo("Col60a89") += nSaldoFac
                    ElseIf nDiasVencido > 29 Then
                        drAnexo("Col30a59") += nSaldoFac
                    ElseIf nDiasVencido > 0 Then
                        drAnexo("Col1a29") += nSaldoFac
                    End If
                    drAnexo("Total") += nSaldoFac

                    ' La siguiente condición permite tener en el campo Retraso el máximo número de
                    ' días vencidos del contrato

                    If nDiasVencido > drAnexo("Retraso") Then
                        drAnexo("Retraso") = nDiasVencido
                    End If

                End If

            End If

        Next

        ' Esta vista de la tabla Clientes me permite que primero aparezcan los clientes que tienen mayor atraso en días

        dvClientes = New DataView(dtClientes)
        dvClientes = dtClientes.DefaultView
        dvClientes.Sort = "Col90omas DESC, Col60a89 DESC, Col30a59 DESC, Col1a29 DESC"

        ''Select Case cbPromotores.SelectedIndex()
        ''    Case 0
        ''        ' Listado General (excepto NAVOJOA)
        ''        dvClientes.RowFilter() = "Sucursal <> '03'"
        ''    Case 1
        ''        ' Cartera Exigible de 1 a 29 días
        ''        dvClientes.RowFilter() = "Sucursal <> '03' AND Retraso >= 1 AND Retraso <= 29"
        ''    Case 2
        ''        ' Cartera Exigible de 30 a 59 días
        ''        dvClientes.RowFilter() = "Sucursal <> '03' AND Retraso >= 30 AND Retraso <= 59"
        ''    Case 3
        ''        ' Cartera Exigible de 60 a 89 días
        ''        dvClientes.RowFilter() = "Sucursal <> '03' AND Retraso >= 60 AND Retraso <= 89"
        ''    Case 4
        ''        ' Cartera Vencida
        ''        dvClientes.RowFilter() = "Sucursal <> '03' AND Retraso >= 90"
        ''    Case 5
        ''        ' Cartera Castigada
        ''        dvClientes.RowFilter() = ""
        ''    Case 6
        ''        ' Cartera Exigible y Vencida de NAVOJOA
        ''        dvClientes.RowFilter() = "Sucursal = '03'"
        ''    Case Is >= 7
        ''        ' Cartera por Promotor
        ''        dvClientes.RowFilter() = "Promotor = '" & cbPromotores.SelectedItem & "'"
        ''End Select

        ' Esta vista de la tabla Anexos me permite ordenar el reporte por antigüedad del saldo

        dvAnexos = New DataView(dtAnexos)
        dvAnexos = dtAnexos.DefaultView
        dvAnexos.Sort = "Col90omas DESC, Col60a89 DESC, Col30a59 DESC, Col1a29 DESC"

        ' Por cada registro de la tabla Clientes, debo crear un filtro de la tabla Anexos dejando visibles
        ' solamente los contratos del cliente en cuestión, añadiéndolos a una tercera tabla llamada Reporte

        nClientes = dvClientes.Count()
        TA.DeleteAll()
        For i = 0 To nClientes - 1
            cCliente = dvClientes.Item(i)("Cliente")
            dvAnexos.RowFilter() = "Cliente = " & cCliente
            nAnexos = dvAnexos.Count()
            For j = 0 To nAnexos - 1

                If dvAnexos.Item(j)("Vencida") <> "C" Then
                    drReporte = dtReporte.NewRow()
                    drReporte("Anexo") = dvAnexos.Item(j)("Anexo")
                    drReporte("Cliente") = dvAnexos.Item(j)("Cliente")
                    drReporte("Nombre") = dvAnexos.Item(j)("Nombre")
                    drReporte("Letra") = dvAnexos.Item(j)("Letra")
                    drReporte("Plazo") = dvAnexos.Item(j)("Plazo")
                    drReporte("Dia") = dvAnexos.Item(j)("Dia")
                    drReporte("Saldo") = dvAnexos.Item(j)("Saldo")
                    drReporte("Col1a29") = dvAnexos.Item(j)("Col1a29")
                    drReporte("Col30a59") = dvAnexos.Item(j)("Col30a59")
                    drReporte("Col60a89") = dvAnexos.Item(j)("Col60a89")
                    drReporte("Col90omas") = dvAnexos.Item(j)("Col90omas")
                    drReporte("Total") = dvAnexos.Item(j)("Total")
                    drReporte("Retraso") = dvAnexos.Item(j)("Retraso")
                    drReporte("Reestructura") = dvAnexos.Item(j)("Reestructura")
                    drReporte("PI") = dvAnexos.Item(j)("PI")
                    drReporte("OC") = dvAnexos.Item(j)("OC")
                    drReporte("Recursos") = dvAnexos.Item(j)("Recursos")
                    drReporte("InicialesPromotor") = dvAnexos.Item(j)("InicialesPromotor")
                    drReporte("Gestion") = dvAnexos.Item(j)("Gestion")
                    drReporte("Seguimiento") = dvAnexos.Item(j)("Seguimiento")
                    drReporte("FechaLlamada") = dvAnexos.Item(j)("FechaLlamada")
                    dtReporte.Rows.Add(drReporte)
                    TA.Update(drReporte)
                    drReporte.AcceptChanges()

                End If
            Next

        Next





        dsAgil.Tables.Remove("Facturas")
        dsAgil.Tables.Remove("Anexos")
        dsAgil.Tables.Remove("PagosIniciales")
        dsAgil.Tables.Remove("Opciones")
        dsAgil.Tables.Remove("Reestructurados")
        dsAgil.Tables.Remove("Seguimientos")

        dsAgil.Tables.Add(dtReporte)

        ' Descomentar la siguiente línea en caso de que se deseara modificar el reporte rptRepAntig
        ' dsAgil.WriteXml("C:\Schema1.xml", XmlWriteMode.WriteSchema)



        ''newrptRepAntig2 = New rptRepAntig
        ''newrptRepAntig2.SetDataSource(dsAgil)
        ''cReportTitle = "REPORTE DE ANTIGUEDAD DE SALDOS AL " & Mes(cFecha)
        ''If cbPromotores.SelectedIndex = 0 Then
        ''    cReportComments = "LISTADO GENERAL (Excepto NAVOJOA)"
        ''ElseIf cbPromotores.SelectedIndex <= 6 Then
        ''    cReportComments = RTrim(cbPromotores.SelectedItem)
        ''Else
        ''    cReportComments = "CLIENTES ASIGNADOS A " & RTrim(cbPromotores.SelectedItem)
        ''End If
        ''newrptRepAntig2.SummaryInfo.ReportTitle = cReportTitle
        ''newrptRepAntig2.SummaryInfo.ReportComments = cReportComments
        ''CrystalReportViewer1.ReportSource = newrptRepAntig2
        ''CrystalReportViewer1.Zoom(100)

        cnAgil.Dispose()
        cm1.Dispose()
        cm2.Dispose()
        cm3.Dispose()
        cm4.Dispose()
        cm5.Dispose()
        cm6.Dispose()

    End Sub

    Sub LlenaHisotriaCartera()
        Dim Cartera As New ProductionDataSetTableAdapters.Vw_CarteraXrecuepararTableAdapter
        Dim Historia As New ProductionDataSetTableAdapters.CarteraXrecuepararHISTTableAdapter
        Dim Cart As New ProductionDataSet.Vw_CarteraXrecuepararDataTable
        Cartera.Fill(Cart)
        For Each r As ProductionDataSet.Vw_CarteraXrecuepararRow In Cart.Rows
            Historia.Insert(r.Promotor, r.Cliente, r.Anexo, r.Letra, r.Plazo, r.Saldo, r.Col1a29, r.Col30a59, r.Col60a89, r.Col90omas, r.Total, r.Retraso, r.Fecha)
        Next
    End Sub

    Sub LeeCSV_Tipo3()
        Dim ta As New ProductionDataSetTableAdapters.FIRA_MOVSTableAdapter
        Dim ds As New ProductionDataSet
        Dim r As ProductionDataSet.FIRA_MOVSRow
        Dim arr As String()
        Dim eRR As String = ""
        Dim Linea As String
        Dim n As Integer
        Dim Nuevo As System.IO.StreamReader
        Dim di1 As New DirectoryInfo(RutaFira)
        Dim f1 As FileInfo() = di1.GetFiles("3*.csv")

        total = f1.Length
        For I As Integer = 0 To f1.Length - 1
            Nuevo = New System.IO.StreamReader(f1(I).FullName, Text.Encoding.GetEncoding(1252))
            n = 0
            While Not Nuevo.EndOfStream
                n += 1
                Linea = Nuevo.ReadLine()
                If n = 1 Then
                    Linea = Nuevo.ReadLine()
                End If
                Linea = Linea.Replace("FINAGIL, S.A", "FINAGIL S.A")
                Linea = Linea.Replace("""", "")
                arr = Linea.Split(",")
                r = ds.FIRA_MOVS.NewFIRA_MOVSRow
                For w As Integer = 0 To arr.Length - 1
                    If arr.Length = 58 And w = 27 Then
                        r(w) = arr(w) & arr(w + 1)
                    ElseIf arr.Length = 58 And w = 57 Then

                    ElseIf arr.Length = 58 And w > 27 Then
                        r(w) = arr(w + 1)
                    Else
                        r(w) = arr(w)
                    End If
                Next
                Try
                    ds.FIRA_MOVS.AddFIRA_MOVSRow(r)
                    ds.FIRA_MOVS.GetChanges()
                    ta.Update(ds.FIRA_MOVS)
                Catch ex As Exception
                    eRR = String.Join("|", arr) & vbCrLf
                End Try
            End While
            If n > 0 Then
                Nuevo.Close()
            End If
            If eRR > "" Then
                EnviaConfirmacion("ecacerest@finagil.com.mx", eRR, "Error en datos Tipo3 FIRA_MOVS")
            End If
            ta.CorrigeFechas()
            Console.WriteLine(porcentaje.ToString("p") & " " & f1(I).Name)
        Next
    End Sub

End Module
