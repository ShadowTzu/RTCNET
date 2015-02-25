Imports System.Reflection

Module Module1
    Private mRTC As New RTCNET
    Sub Main()
        'Compile dotnet code
        'Example.vb is copied to executable folder. Edit it with a text editor and see the result in real time!
        mRTC.Compile_FromFile("myScript", IO.Path.Combine(My.Application.Info.DirectoryPath, "Example.vb"))

        'Create Class Example with name "myClass"
        mRTC.Create_Instance("myScript", "Example", "myClass")

        Dim last_error As String = String.Empty
        Do While True
            'If errors are found then write only one time in the console
            If mRTC.HasError("myScript") Then
                If last_error <> mRTC.Error_Message("myScript") Then
                    last_error = mRTC.Error_Message("myScript")
                    Console.WriteLine(last_error)
                End If
            Else
                last_error = String.Empty
                'Call sub Test
                mRTC.Call_Function("myScript", "myClass", "Test")

                'Call Sub Write_Text with args: This is a text!
                mRTC.Call_Function("myScript", "myClass", "Write_Text", "This is a text!")

                'Call function Add with args 123 and 456 and write the result in the console
                Console.WriteLine(mRTC.Call_Function("myScript", "myClass", "Add", 123, 456))
            End If

        Loop

        mRTC.Dispose()
    End Sub

End Module
