Attribute VB_Name = "Usage"
Option Explicit

' Uses the peer addin - required for the peer reference to persist on save
' (VBA drops unused project references).
Public Function GetGreeting() As String
    GetGreeting = AddinPeer.SayHello()
End Function
