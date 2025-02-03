VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} lblPrime 
   Caption         =   "Prime d'assurances"
   ClientHeight    =   4725
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   9898.001
   OleObjectBlob   =   "Userform1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "lblPrime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCalculer_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub Label1_Click()

End Sub

Private Sub btnCalculer_Click()
    Dim age As Integer
    Dim nbSinistres As Long
    Dim prime As Double
    Dim prenom As String
    Dim nom As String
    
     ' Récupération des valeurs pour le prénom et le nom
    prenom = txtPrenom.Value
    nom = txtNom.Value
    
    
    ' Vérification prénom et nom et conversion des saisies'
     
    If Trim(prenom) = "" Or IsNumeric(prenom) Then
        MsgBox "Veuillez saisir un prénom valide (du texte et non un nombre).", vbExclamation
        Exit Sub
    End If
    
    
    If Trim(nom) = "" Or IsNumeric(nom) Then
        MsgBox "Veuillez saisir un nom valide (du texte et non un nombre).", vbExclamation
        Exit Sub
    End If
    
    
    If IsNumeric(txtAge.Value) And IsNumeric(txtSinistres.Value) Then
        age = CInt(txtAge.Value)
        nbSinistres = CLng(txtSinistres.Value)
        
    Else
        MsgBox "Veuillez entrer des valeurs numériques pour l'âge et le nombre de sinistres.", vbExclamation
        Exit Sub
    End If
    
    ' Vérifier si l'âge est inférieur à 18 ans
    If age < 18 Then
        MsgBox "L'âge doit être supérieur ou égal à 18 ans.", vbCritical
        Exit Sub
    
    End If
    
    ' Calcul de la prime
    prime = CalculerPrime(age, nbSinistres)
    
  ' Vérifier la valeur de prime pour le débogage
    ' MsgBox "Prime calculée : " & prime

    ' Affichage du résultat dans le label du UserForm
    lblContrat.Caption = "Contrat d'assurance : " & prenom & " " & nom & ""
    lblResultat.Caption = "Bonjour " & prenom & " " & nom & " votre prime d'assurance est de " & Format(prime, "0.00") & " €"
End Sub
    


Private Sub lblAge_Click()

End Sub

Private Sub LblResultat_Click()

End Sub

Private Sub lblSinistres_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub txtAge_Change()

End Sub

Private Sub UserForm_Click()

End Sub
