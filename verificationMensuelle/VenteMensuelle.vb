Public Class VenteMensuelle
    Private nbBoitesVendues As Integer
    Private prixHT As Decimal
    Private prixLabo As Decimal

    '----------------------------------------------------------------
    ' Constructeur de la classe
    '----------------------------------------------------------------
    Sub New()
        Me.nbBoitesVendues = 0
        Me.prixHT = 0.0
        Me.prixLabo = 0.0
    End Sub

    '----------------------------------------------------------------
    ' Accesseurs
    '----------------------------------------------------------------
    Public Function getNbBoitesVendues() As Integer
        Return Me.nbBoitesVendues
    End Function

    Public Sub setNbBoitesVendues(ByVal unNombre As Integer)
        'If (unNombre < 0) Then
        '    Throw New ApplicationException("le nombre de boîtes vendues doit être positif.")
        'End If
        Me.nbBoitesVendues = unNombre
    End Sub

    Public Function getPrixHT() As Decimal
        Return Me.prixHT
    End Function

    Public Sub setPrixHT(ByVal unPrixHT As Decimal)
        'If (unPrixHT < 0) Then
        '    Throw New ApplicationException("le prix hors taxes doit être positif.")
        'End If
        Me.prixHT = unPrixHT
    End Sub

    Public Function getPrixLabo() As Decimal
        Return Me.prixLabo
    End Function

    Public Sub setPrixLabo(ByVal unPrixLabo As Decimal)
        'If (unPrixLabo < 0) Then
        '    Throw New ApplicationException("Le prix Laboratoire doit être positif.")
        'End If
        Me.prixLabo = unPrixLabo
    End Sub

End Class
