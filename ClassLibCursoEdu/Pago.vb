Imports System.Data.SqlClient

Public Class Pago
    Private mPagoUsuarioId As Integer
    Private mPagoCursoId As Integer
    Private mPagoFecha As DateTime
    Private mPagoDescuento As Integer

    Public Property PagoUsuarioId As Integer
        Get
            Return mPagoUsuarioId
        End Get
        Set(value As Integer)
            mPagoUsuarioId = value
        End Set
    End Property

    Public Property PagoCursoId As Integer
        Get
            Return mPagoCursoId
        End Get
        Set(value As Integer)
            mPagoCursoId = value
        End Set
    End Property

    Public Property PagoFecha As DateTime
        Get
            Return mPagoFecha
        End Get
        Set(value As DateTime)
            mPagoFecha = value
        End Set
    End Property

    Public Property PagoDescuento As Integer
        Get
            Return mPagoDescuento
        End Get
        Set(value As Integer)
            mPagoDescuento = value
        End Set
    End Property

    Public Function GuardarPago() As Boolean
        Using cnx As New SqlConnection(ModuleDB.server)
            Dim cmd As New SqlCommand("dbo.pagoAlta", cnx)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@pago_usuario_id", PagoUsuarioId)
            cmd.Parameters.AddWithValue("@pago_curso_id", PagoCursoId)
            cmd.Parameters.AddWithValue("@pago_fecha", PagoFecha)
            cmd.Parameters.AddWithValue("@pago_descuento", PagoDescuento)

            Try
                cnx.Open()
                cmd.ExecuteNonQuery()
                Return True
            Catch ex As Exception
                Console.WriteLine("Error al guardar el pago: " & ex.Message)
                Return False
            Finally
                If cnx.State = ConnectionState.Open Then
                    cnx.Close()
                End If
            End Try
        End Using
    End Function

    Public Function UsuarioHaPagadoCurso(usuarioId As Integer, cursoId As Integer) As Boolean
        Using cnx As New SqlConnection(ModuleDB.server)
            Dim cmd As New SqlCommand("SELECT COUNT(*) FROM pago WHERE pago_usuario_id = @usuarioId AND pago_curso_id = @cursoId;", cnx)
            cmd.Parameters.AddWithValue("@usuarioId", usuarioId)
            cmd.Parameters.AddWithValue("@cursoId", cursoId)

            Try
                cnx.Open()
                Dim rowCount As Integer = CInt(cmd.ExecuteScalar())
                Return rowCount > 0
            Catch ex As Exception
                Console.WriteLine("Error al verificar si el usuario ha pagado el curso: " & ex.Message)
                Return False
            Finally
                If cnx.State = ConnectionState.Open Then
                    cnx.Close()
                End If
            End Try
        End Using
    End Function

End Class
