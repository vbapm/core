Attribute VB_Name = "Bonjour"
Option Explicit

' Français : é è ê ë à â ä ù û ü ç É È Ê Ë À Â Ä Ù Û Ü Ç
' Allemand : ä ö ü ß Ä Ö Ü
' Espagnol : á é í ó ú ñ ¿ ¡

Public Function Hello() As String
    Hello = "Voilà les caractères accentués !"
End Function