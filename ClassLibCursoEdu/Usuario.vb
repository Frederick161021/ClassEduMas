Imports System.Data.SqlClient

Public Class Usuario
    Private MUsuarioId As Integer
    Private MUsuarioRolId As Integer
    Private MUsuarioNombre As String
    Private MUsuarioContrasena As String
    Private MUsuarioNumTarjeta As Integer
    Private MUsuarioCVV As Integer
    Private MUsuarioEstado As Integer

    Public Property UsuarioId() As Integer
        Get
            Return MUsuarioId
        End Get
        Set(value As Integer)
            MUsuarioId = value
        End Set
    End Property

    Public Property UsuarioRolId() As Integer
        Get
            Return MUsuarioRolId
        End Get
        Set(value As Integer)
            MUsuarioRolId = value
        End Set
    End Property

    Public Property UsuarioNombre() As String
        Get
            Return MUsuarioNombre
        End Get
        Set(value As String)
            MUsuarioNombre = value
        End Set
    End Property

    Public Property UsuarioContrasena() As String
        Get
            Return MUsuarioContrasena
        End Get
        Set(value As String)
            MUsuarioContrasena = value
        End Set
    End Property

    Public Property UsuarioNumTarjeta() As Integer
        Get
            Return MUsuarioNumTarjeta
        End Get
        Set(value As Integer)
            MUsuarioNumTarjeta = value
        End Set
    End Property

    Public Property UsuarioCVV() As Integer
        Get
            Return MUsuarioCVV
        End Get
        Set(value As Integer)
            MUsuarioCVV = value
        End Set
    End Property

    Public Property UsuarioEstado() As Integer
        Get
            Return MUsuarioEstado
        End Get
        Set(value As Integer)
            MUsuarioEstado = value
        End Set
    End Property

    Public Function UsuarioAlta() As Boolean
        Dim cnx As New SqlConnection(ModuleDB.server)
        Dim cmd As New SqlCommand("dbo.usuarioAlta", cnx)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@usuario_rol_id", MUsuarioRolId))
        cmd.Parameters.Add(New SqlParameter("@usuario_nombre", MUsuarioNombre))
        cmd.Parameters.Add(New SqlParameter("@usuario_contraseña", MUsuarioContrasena))
        cmd.Parameters.Add(New SqlParameter("@usuario_num_tarjeta", MUsuarioNumTarjeta))
        cmd.Parameters.Add(New SqlParameter("@usuario_cvv", MUsuarioCVV))
        cmd.Parameters.Add(New SqlParameter("@usuario_estado", MUsuarioEstado))
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

    Public Function UsuarioBaja() As Boolean
        Dim cnx As New SqlConnection(ModuleDB.server)
        Dim cmd As New SqlCommand("dbo.usuarioBaja", cnx)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@usuario_id", MUsuarioId))
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

    Public Function UsuarioActualiza() As Boolean
        Dim cnx As New SqlConnection(ModuleDB.server)
        Dim cmd As New SqlCommand("dbo.usuarioActualiza", cnx)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@usuario_id", MUsuarioId))
        cmd.Parameters.Add(New SqlParameter("@usuario_rol_id", MUsuarioRolId))
        cmd.Parameters.Add(New SqlParameter("@usuario_nombre", MUsuarioNombre))
        cmd.Parameters.Add(New SqlParameter("@usuario_contraseña", MUsuarioContrasena))
        cmd.Parameters.Add(New SqlParameter("@usuario_num_tarjeta", MUsuarioNumTarjeta))
        cmd.Parameters.Add(New SqlParameter("@usuario_cvv", MUsuarioCVV))
        cmd.Parameters.Add(New SqlParameter("@usuario_estado", MUsuarioEstado))
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

    Public Function UsuarioConsulta() As Boolean
        Dim cnx As New SqlConnection(ModuleDB.server)
        Dim cmd As New SqlCommand("dbo.usuarioConsulta", cnx)
        cmd.CommandType = CommandType.StoredProcedure
        Dim usuarioNombre1 As String = ""
        Dim usuarioContrasena1 As String = ""
        Dim usuarioNumTarjeta1 As Integer
        Dim usuarioCVV1 As Integer
        Dim usuarioEstado1 As Integer
        Dim pasar As Boolean = False
        cmd.Parameters.Add(New SqlParameter("@usuario_id", MUsuarioId))
        cnx.Open()
        Dim leer As SqlDataReader
        leer = cmd.ExecuteReader
        If leer.Read() Then
            usuarioNombre1 = leer("usuario_nombre").ToString()
            usuarioContrasena1 = leer("usuario_contraseña").ToString()
            usuarioNumTarjeta1 = Convert.ToInt32(leer("usuario_num_tarjeta"))
            usuarioCVV1 = Convert.ToInt32(leer("usuario_cvv"))
            usuarioEstado1 = Convert.ToInt32(leer("usuario_estado"))
            UsuarioNombre = usuarioNombre1
            UsuarioContrasena = usuarioContrasena1
            UsuarioNumTarjeta = usuarioNumTarjeta1
            UsuarioCVV = usuarioCVV1
            UsuarioEstado = usuarioEstado1
            cnx.Close()
        End If
        If pasar Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function UsuarioLogin() As Boolean
        Dim cnx As New SqlConnection(ModuleDB.server)
        Dim cmd As New SqlCommand("dbo.usuarioLogin", cnx)
        cmd.CommandType = CommandType.StoredProcedure
        Dim usuarioId1 As Integer
        Dim usuario_rol_id1 As Integer
        Dim usuarioNombre1 As String = ""
        Dim usuarioContrasena1 As String = ""
        Dim usuarioNumTarjeta1 As Integer
        Dim usuarioCVV1 As Integer
        Dim usuarioEstado1 As Integer
        Dim pasar As Boolean = False
        cmd.Parameters.Add(New SqlParameter("@usuario_nombre", MUsuarioNombre))
        cmd.Parameters.Add(New SqlParameter("@usuario_contraseña", MUsuarioContrasena))
        cnx.Open()
        Dim leer As SqlDataReader
        leer = cmd.ExecuteReader
        If leer.Read() Then
            usuarioId1 = Convert.ToInt32(leer("usuario_id"))
            usuario_rol_id1 = Convert.ToInt32(leer("usuario_rol_id"))
            usuarioNombre1 = MUsuarioNombre
            usuarioContrasena1 = MUsuarioContrasena
            usuarioNumTarjeta1 = Convert.ToInt32(leer("usuario_num_tarjeta"))
            usuarioCVV1 = Convert.ToInt32(leer("usuario_cvv"))
            usuarioEstado1 = Convert.ToInt32(leer("usuario_estado"))
            UsuarioId = usuarioId1
            MUsuarioRolId = usuario_rol_id1
            UsuarioNombre = usuarioNombre1
            UsuarioContrasena = usuarioContrasena1
            UsuarioNumTarjeta = usuarioNumTarjeta1
            UsuarioCVV = usuarioCVV1
            UsuarioEstado = usuarioEstado1
            cnx.Close()
        End If
        If pasar Then
            Return False
        Else
            Return True
        End If
    End Function

End Class
