Module Modules
    Function RecupEntreCar(ByVal LaChaine As String, ByVal BaliseOuverte As String, ByVal BaliseFermee As String) As String
        Dim Resultat, PartieD, PartieG, rTab As String

        PartieD = Droite_de(LaChaine, BaliseOuverte)
        PartieG = Gauche_de(PartieD, BaliseFermee)
        Resultat = PartieG
        Return Resultat
    End Function

    Function Gauche_de(ByVal LaChaine As String, ByVal Caract As String)
        Dim resultat, CarEnCours As String
        Dim i, n As Integer

        While CarEnCours <> Caract
            i = i + 1
            CarEnCours = Mid(LaChaine, i, 1)
        End While

        resultat = Mid(LaChaine, 1, i - 1)

        Return resultat
    End Function


    Function Droite_de(ByVal LaChaine As String, ByVal Caract As String)
        Dim resultat, CarEnCours As String
        Dim i, n As Integer

        While CarEnCours <> Caract
            i = i + 1
            CarEnCours = Mid(LaChaine, Len(LaChaine) - i, 1)
        End While

        resultat = Right(LaChaine, i)

        Return resultat
    End Function

End Module
