Attribute VB_Name = "Module1"
Option Explicit

' Fonction de calcul de la prime d'assurance
Public Function CalculerPrime(age As Integer, nbSinistres As Long) As Double
    Dim primeBase As Double
    primeBase = 100  ' Prime de base

    ' Majorations en fonction de l'âge
    If age < 25 Then
        primeBase = primeBase * 1.5
    ElseIf age >= 25 And age <= 65 Then
        primeBase = primeBase * 1.2
    Else
        primeBase = primeBase * 1.3
    End If

    ' Ajout d'une majoration pour chaque sinistre
    If nbSinistres > 0 Then
        primeBase = primeBase + (nbSinistres * 50)
    End If

    CalculerPrime = primeBase
End Function

