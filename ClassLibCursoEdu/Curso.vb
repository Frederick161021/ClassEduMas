Imports System.Data.SqlClient
Public Class Curso
    Private MCursoID As Integer
    Private MCursoNombre As Integer
    Private MCursoDescripcion As String
    Private MCursoCosto As String

    Public Property CursoID() As Integer
        Get
            Return MCursoID
        End Get
        Set(value As Integer)
            MCursoID = value
        End Set
    End Property

    Public Property CursoNombre() As String
        Get
            Return MCursoNombre
        End Get
        Set(value As String)
            MCursoNombre = value
        End Set
    End Property

    Public Property CursoDescripcion() As String
        Get
            Return MCursoDescripcion
        End Get
        Set(value As String)
            MCursoDescripcion = value
        End Set
    End Property

    Public Property CursoCosto() As Integer
        Get
            Return MCursoCosto
        End Get
        Set(value As Integer)
            MCursoCosto = value
        End Set
    End Property


    Public Function CursoAlta() As Boolean
        Dim cnx As New SqlConnection(ModuleDB.server)
        Dim cmd As New SqlCommand("dbo.CursoAlta", cnx)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@curso_nombre", MCursoNombre))
        cmd.Parameters.Add(New SqlParameter("@curso_descripcion", MCursoDescripcion))
        cmd.Parameters.Add(New SqlParameter("@curso_cost", MCursoCosto))
        Try
            cnx.Open()
            cmd.ExecuteScalar()
            Return True
        Catch ex As Exception
            ' Manejo de excepciones
        Finally
            If cnx.State = ConnectionState.Open Then
                cnx.Close()
            End If
        End Try
        Return False
    End Function

    Public Function CursoBaja() As Boolean
        Dim cnx As New SqlConnection(ModuleDB.server)
        Dim cmd As New SqlCommand("dbo.CursoBaja", cnx)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@curso_id", MCursoID))
        Try
            cnx.Open()
            cmd.ExecuteScalar()
            Return True
        Catch ex As Exception
            ' Manejo de excepciones
        Finally
            If cnx.State = ConnectionState.Open Then
                cnx.Close()
            End If
        End Try
        Return False
    End Function

    Public Function CursoActualiza() As Boolean
        Dim cnx As New SqlConnection(ModuleDB.server)
        Dim cmd As New SqlCommand("dbo.CursoActualiza", cnx)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@curso_id", MCursoID))
        cmd.Parameters.Add(New SqlParameter("@curso_nombre", MCursoNombre))
        cmd.Parameters.Add(New SqlParameter("@curso_descripcion", MCursoDescripcion))
        cmd.Parameters.Add(New SqlParameter("@curso_cost", MCursoCosto))
        Try
            cnx.Open()
            cmd.ExecuteScalar()
            Return True
        Catch ex As Exception
            ' Manejo de excepciones
        Finally
            If cnx.State = ConnectionState.Open Then
                cnx.Close()
            End If
        End Try
        Return False
    End Function

    Public Function CursoConsulta() As Boolean
        Dim cnx As New SqlConnection(ModuleDB.server)
        Dim cmd As New SqlCommand("dbo.CursoConsulta", cnx)
        cmd.CommandType = CommandType.StoredProcedure
        Dim CursoNombre1 As String = ""
        Dim CursoDescripcion1 As String = ""
        Dim CursoCosto1 As Integer
        Dim pasar As Boolean = False
        cmd.Parameters.Add(New SqlParameter("@curso_id", MCursoID))
        cnx.Open()
        Dim leer As SqlDataReader
        leer = cmd.ExecuteReader
        If leer.Read() Then
            CursoNombre1 = leer("curso_nombre").ToString()
            CursoDescripcion1 = leer("curso_descripcion").ToString()
            CursoCosto1 = Convert.ToInt32(leer("curso_cost"))
            CursoNombre = CursoNombre1
            CursoDescripcion = CursoDescripcion1
            CursoCosto = CursoCosto1
            cnx.Close()
        End If
        If pasar Then
            Return False
        Else
            Return True
        End If
    End Function



End Class
