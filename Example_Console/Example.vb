Imports System

Public Class Example
    Public Sub Test()
        Console.WriteLine("Test!")
    End Sub

    Public Sub Write_Text(Text As String)
        Console.WriteLine(Text)
    End Sub

    Public Function Add(A As Single, B As Single) As Single
        Return (A + B)
    End Function
End Class
