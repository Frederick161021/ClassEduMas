Imports System.Data.SqlClient

Public Class Contenido
    Private MContenidoId As Integer
    Private MContenidoCursoId As Integer
    Private MContenidoModulo As Integer
    Private MContenidoNivel As Integer
    Private MContenidoTitulo As String
    Private MContenidoTexto As String
    Private MContenidoImagen As String

    Public Property ContenidoId() As Integer
        Get
            Return MContenidoId
        End Get
        Set(value As Integer)
            MContenidoId = value
        End Set
    End Property

    Public Property ContenidoCursoId() As Integer
        Get
            Return MContenidoCursoId
        End Get
        Set(value As Integer)
            MContenidoCursoId = value
        End Set
    End Property

    Public Property ContenidoModulo() As Integer
        Get
            Return MContenidoModulo
        End Get
        Set(value As Integer)
            MContenidoModulo = value
        End Set
    End Property

    Public Property ContenidoNivel() As Integer
        Get
            Return MContenidoNivel
        End Get
        Set(value As Integer)
            MContenidoNivel = value
        End Set
    End Property

    Public Property ContenidoTitulo() As String
        Get
            Return MContenidoTitulo
        End Get
        Set(value As String)
            MContenidoTitulo = value
        End Set
    End Property

    Public Property ContenidoTexto() As String
        Get
            Return MContenidoTexto
        End Get
        Set(value As String)
            MContenidoTexto = value
        End Set
    End Property

    Public Property ContenidoImagen() As String
        Get
            Return MContenidoImagen
        End Get
        Set(value As String)
            MContenidoImagen = value
        End Set
    End Property

    Public Function ContenidoAlta() As Boolean
        Dim cnx As New SqlConnection(ModuleDB.server)
        Dim cmd As New SqlCommand("dbo.contenidoAlta", cnx)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@contenido_id", MContenidoId))
        cmd.Parameters.Add(New SqlParameter("@contenido_curso_id", MContenidoCursoId))
        cmd.Parameters.Add(New SqlParameter("@contenido_modulo", MContenidoModulo))
        cmd.Parameters.Add(New SqlParameter("@contenido_nivel", MContenidoNivel))
        cmd.Parameters.Add(New SqlParameter("@contenido_titulo", MContenidoTitulo))
        cmd.Parameters.Add(New SqlParameter("@contenido_texto", MContenidoTexto))
        cmd.Parameters.Add(New SqlParameter("@contenido_imagen", MContenidoImagen))
        Try
            cnx.Open()
            cmd.ExecuteNonQuery()
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

    Public Function ContenidoBaja() As Boolean
        Dim cnx As New SqlConnection(ModuleDB.server)
        Dim cmd As New SqlCommand("dbo.contenidoBaja", cnx)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@contenido_id", MContenidoId))
        Try
            cnx.Open()
            cmd.ExecuteNonQuery()
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

    Public Function ContenidoActualiza() As Boolean
        Dim cnx As New SqlConnection(ModuleDB.server)
        Dim cmd As New SqlCommand("dbo.contenidoActualiza", cnx)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@contenido_id", MContenidoId))
        cmd.Parameters.Add(New SqlParameter("@contenido_curso_id", MContenidoCursoId))
        cmd.Parameters.Add(New SqlParameter("@contenido_modulo", MContenidoModulo))
        cmd.Parameters.Add(New SqlParameter("@contenido_nivel", MContenidoNivel))
        cmd.Parameters.Add(New SqlParameter("@contenido_titulo", MContenidoTitulo))
        cmd.Parameters.Add(New SqlParameter("@contenido_texto", MContenidoTexto))
        cmd.Parameters.Add(New SqlParameter("@contenido_imagen", MContenidoImagen))
        Try
            cnx.Open()
            cmd.ExecuteNonQuery()
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

    Public Function ContenidoConsulta() As Boolean
        Dim cnx As New SqlConnection(ModuleDB.server)
        Dim cmd As New SqlCommand("dbo.contenidoConsulta", cnx)
        cmd.CommandType = CommandType.StoredProcedure
        Dim contenidoTitulo1, contenidoTexto1, contenidoImagen1 As String
        Dim pasar As Boolean
        cmd.Parameters.Add(New SqlParameter("@contenido_id", MContenidoId))
        cnx.Open()
        Dim leer As SqlDataReader
        leer = cmd.ExecuteReader
        If leer.Read() Then
            contenidoTitulo1 = leer(4).ToString
            contenidoTexto1 = leer(5).ToString
            contenidoImagen1 = leer(6).ToString
            ContenidoTitulo = contenidoTitulo1
            ContenidoTexto = contenidoTexto1
            ContenidoImagen = contenidoImagen1
            cnx.Close()
        End If
        If pasar Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function ConenidoCargar() As Boolean
        Dim cnx As New SqlConnection(ModuleDB.server)
        Dim cmd As New SqlCommand("dbo.contenidoCargar", cnx)
        cmd.CommandType = CommandType.StoredProcedure
        Dim pasar As Boolean = False
        cmd.Parameters.Add(New SqlParameter("@contenido_curso_id", MContenidoCursoId))
        cmd.Parameters.Add(New SqlParameter("@contenido_modulo", MContenidoModulo))
        cnx.Open()
        Dim leer As SqlDataReader
        leer = cmd.ExecuteReader
        If leer.Read() Then
            MContenidoId = Convert.ToInt32(leer("contenido_id"))
            MContenidoNivel = Convert.ToInt32(leer("contenido_nivel"))
            MContenidoTitulo = leer("contenido_titulo").ToString
            MContenidoTexto = leer("contenido_texto").ToString
            MContenidoImagen = leer("contenido_imagen").ToString
            cnx.Close()
        End If
        If pasar Then
            Return False
        Else
            Return True
        End If
    End Function

End Class
