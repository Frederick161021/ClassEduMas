Imports System.Data.SqlClient

Public Class Actividad
    Private MActividadUsuarioId As Integer
    Private MActividadFecha As DateTime

    Public Property ActividadUsuarioId As Integer
        Get
            Return MActividadUsuarioId
        End Get
        Set(value As Integer)
            MActividadUsuarioId = value
        End Set
    End Property

    Public Property ActividadFecha As DateTime
        Get
            Return MActividadFecha
        End Get
        Set(value As DateTime)
            MActividadFecha = value
        End Set
    End Property

    Public Function ActividadAlta() As Boolean
        Using cnx As New SqlConnection(ModuleDB.server)
            Dim cmd As New SqlCommand("dbo.actvidadAlta", cnx)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@actividad_usuario_id", MActividadUsuarioId))
            cmd.Parameters.Add(New SqlParameter("@actividad_fecha", MActividadFecha))
            Try
                cnx.Open()
                cmd.ExecuteNonQuery()
                Return True
            Catch ex As Exception
                Console.WriteLine("Error al registrar la actividad: " & ex.Message)
                Return False
            Finally
                If cnx.State = ConnectionState.Open Then
                    cnx.Close()
                End If
            End Try
        End Using
    End Function
End Class
