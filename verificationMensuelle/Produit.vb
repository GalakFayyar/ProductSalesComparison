Public Class Produit

    Private codeCIP As String
    Private designation As String
    Private laboratoire As String
    Private codeTVA As Integer

    Private ventes As Hashtable

    '----------------------------------------------------------------
    ' Constructeur de la classe
    '----------------------------------------------------------------
    Sub New(ByVal unCode As String)
        Me.codeCIP = unCode
            Me.designation = "Pas de désignation"
            Me.laboratoire = "Pas de laboratoire"
            Me.codeTVA = 0
            Me.ventes = New Hashtable
    End Sub

    '----------------------------------------------------------------
    ' Ajout de Ventes
    '----------------------------------------------------------------
    Public Sub ajouterVenteMensuelle(ByVal unMois As Integer, ByVal uneVente As VenteMensuelle)
        If (Me.ventes.ContainsKey(unMois)) Then
            Throw New ApplicationException("l'identifiant '" & unMois & "' existe déja dans la table.")
        End If
        Me.ventes.Add(unMois, uneVente)
    End Sub

    '----------------------------------------------------------------
    ' Accesseurs
    '----------------------------------------------------------------
    Public Function getCodeCIP() As String
        Return Me.codeCIP
    End Function

    Public Function getDesignation() As String
        Return Me.designation
    End Function

    Public Sub setDesignation(ByVal uneDesignation As String)
        Me.designation = uneDesignation
    End Sub

    Public Function getLaboratoire() As String
        Return Me.laboratoire
    End Function

    Public Sub setLaboratoire(ByVal unLaboratoire As String)
        Me.laboratoire = unLaboratoire
    End Sub

    Public Function getCodeTVA() As Integer
        Return Me.codeTVA
    End Function

    Public Sub setCodeTVA(ByVal unCodeTVA As Integer)
        Me.codeTVA = unCodeTVA
    End Sub

    Public Function getVentes() As Hashtable
        Return ventes
    End Function

End Class
