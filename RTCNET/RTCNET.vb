Imports Microsoft.VisualBasic
Imports System
Imports System.Text
Imports System.CodeDom.Compiler
Imports System.Reflection
Imports System.IO

Public NotInheritable Class RTCNET
    Implements IDisposable
    Private Structure struct_Data
        Public Assembly As Assembly
        Public IsFile As Boolean
        Public Filename As String
        Public [Error] As String
        Public FileWatcher As IO.FileSystemWatcher
        Public Instances As Dictionary(Of String, struct_Instance)
    End Structure

    Private Structure struct_Instance
        Public [Type] As Type
        Public instance As Object
        Public ClassName As String
        Sub New(mType As Type, mInstance As Object)
            Me.Type = mType
            Me.instance = mInstance
        End Sub
    End Structure

    Private provider As CodeDomProvider
    Private Assemblies As Dictionary(Of String, struct_Data)

    Private Compiler As VBCodeProvider
    Private params As CompilerParameters
    Private FileAssembly As Dictionary(Of String, String)

    Public Sub New()
        provider = New Microsoft.VisualBasic.VBCodeProvider
        Assemblies = New Dictionary(Of String, struct_Data)
        Compiler = New VBCodeProvider
        params = New CompilerParameters
        FileAssembly = New Dictionary(Of String, String)

    End Sub

    ''' <summary>
    ''' Complie Code from a file
    ''' </summary>
    ''' <param name="Assembly_Name">Give a name for your assembly</param>
    ''' <param name="Filename">Path to your filename</param>
    ''' <remarks></remarks>
    Public Sub Compile_FromFile(Assembly_Name As String, Filename As String)
        If Not File.Exists(Filename) Then Return

        Using File As TextReader = IO.File.OpenText(Filename)
            Compile_FromText(Assembly_Name, File.ReadToEnd)
        End Using

        Dim data As struct_Data = Assemblies(Assembly_Name)

        Dim ext As String = Path.GetExtension(Filename)
        data.FileWatcher = New IO.FileSystemWatcher()

        With data.FileWatcher
            .Path = Path.GetDirectoryName(Filename)
            .EnableRaisingEvents = True         
            .Filter = "*" & ext
            .NotifyFilter = (NotifyFilters.LastWrite)
        End With
        AddHandler data.FileWatcher.Changed, AddressOf OnChanged
        data.IsFile = True
        data.Filename = Filename

        Assemblies(Assembly_Name) = data
        FileAssembly.Add(Filename, Assembly_Name)
    End Sub


    ''' <summary>
    ''' Compile from a string
    ''' </summary>
    ''' <param name="Assembly_Name">Give a name for your assembly</param>
    ''' <param name="Source">Source code</param>
    ''' <remarks></remarks>
    Public Sub Compile_FromText(Assembly_Name As String, ByVal Source As String)
        If Assemblies.ContainsKey(Assembly_Name) Then
            If Not (Assemblies(Assembly_Name).Error = String.Empty) Then
                Assemblies.Remove(Assembly_Name)
            Else
                Return
            End If
        End If

        With params
            .GenerateExecutable = False
            .GenerateInMemory = True
            .IncludeDebugInformation = False

            For Each assembly As Assembly In AppDomain.CurrentDomain.GetAssemblies()
                Try
                    Dim location As String = assembly.Location
                    If (Not String.IsNullOrEmpty(location)) AndAlso (Not .ReferencedAssemblies.Contains(location)) Then
                        .ReferencedAssemblies.Add(location)
                    End If
                Catch generatedExceptionName As NotSupportedException
                End Try
            Next
        End With

        Dim results As CompilerResults

        results = Compiler.CompileAssemblyFromSource(params, Source)

        If results.Errors.HasErrors Then
            Dim errors As New StringBuilder("Compiler Errors: ")
            For Each [error] As CompilerError In results.Errors
                errors.AppendFormat("Line {0},{1}: {2}" & vbLf, [error].Line, [error].Column, [error].ErrorText)
                '  Throw New Exception(errors.ToString())
            Next

            Dim struct As struct_Data = Nothing
            struct.Error = errors.ToString
            struct.Instances = New Dictionary(Of String, struct_Instance)
            Assemblies.Add(Assembly_Name, struct)
        Else
            Dim struct As struct_Data
            struct.Assembly = results.CompiledAssembly
            struct.Instances = New Dictionary(Of String, struct_Instance)
            struct.Error = String.Empty
            struct.Filename = String.Empty
            struct.FileWatcher = Nothing
            Assemblies.Add(Assembly_Name, struct)
            '   AppDomain.CurrentDomain.Load(results.CompiledAssembly.GetName)
        End If

    End Sub

    ''' <summary>
    ''' True if errors has found
    ''' </summary>
    ''' <param name="Assembly_Name">Name of your assembly</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function HasError(Assembly_Name As String) As Boolean
        If Not Assemblies.ContainsKey(Assembly_Name) OrElse (Assemblies(Assembly_Name).Error = String.Empty) Then Return False
        Return True
    End Function

    ''' <summary>
    ''' Error message if errors has found
    ''' </summary>
    ''' <param name="Assembly_Name">Name of your assembly</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Error_Message(Assembly_Name As String) As String
        If Not Assemblies.ContainsKey(Assembly_Name) Then Return String.Empty
        Return Assemblies(Assembly_Name).Error
    End Function

    'TODO: OnChanged is called two time (Why?!), I try to protect with IsCompiling but it's not pefect
    Private IsCompiling As Boolean
    Private Sub ReCompile(Assembly_Name As String)
        If IsCompiling = True Then Exit Sub
        IsCompiling = True
        Dim data As struct_Data = Assemblies(Assembly_Name)
        If data.IsFile = False Then Return

        Dim results As CompilerResults
        Dim File As TextReader = Nothing
        Dim CodeString As String = String.Empty
        Dim FileUsed As Boolean = False
        Do
            Try
                FileUsed = False
                File = IO.File.OpenText(data.Filename)
                CodeString = File.ReadToEnd
            Catch ex As Exception
                FileUsed = True
            End Try
        Loop While FileUsed = True

        results = Compiler.CompileAssemblyFromSource(params, CodeString)

        File.Dispose()
        File = Nothing

        If results.Errors.HasErrors Then
            Dim errors As New StringBuilder("Compiler Errors: ")
            For Each [error] As CompilerError In results.Errors
                errors.AppendFormat("Line {0},{1}: {2}" & vbLf, [error].Line, [error].Column, [error].ErrorText)
            Next
            data.Error = errors.ToString
        Else

            With data
                .Assembly = results.CompiledAssembly
                .Error = String.Empty
            End With
            Dim myInstance As struct_Instance
            For i As Integer = 0 To data.Instances.Count - 1
                myInstance = data.Instances(data.Instances.Keys(i))
                myInstance.instance = data.Assembly.CreateInstance(myInstance.ClassName)
                myInstance.Type = myInstance.instance.GetType
                data.Instances(data.Instances.Keys(i)) = myInstance
            Next
        End If

        Assemblies(Assembly_Name) = data
        IsCompiling = False
    End Sub

    ''' <summary>
    ''' Grab Assembly
    ''' </summary>
    ''' <param name="Assembly_Name">Name of your assembly</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Get_Assembly(Assembly_Name As String) As Assembly
        Do While IsCompiling = True
        Loop
        If Not Assemblies.ContainsKey(Assembly_Name) OrElse (Not Assemblies(Assembly_Name).Error = String.Empty) Then Return Nothing
        Return Assemblies(Assembly_Name).Assembly
    End Function

    ''' <summary>
    ''' Create instance 
    ''' </summary>
    ''' <param name="Assembly_Name">Name of your assembly</param>
    ''' <param name="Class_Name">Name of your class in your source code</param>
    ''' <param name="MyClass_Name">Give a name to this new class</param>
    ''' <remarks></remarks>
    Public Sub Create_Instance(Assembly_Name As String, Class_Name As String, MyClass_Name As String)
        Do While IsCompiling = True
        Loop
        If Not Assemblies.ContainsKey(Assembly_Name) Then Return

        Dim struct_asm As struct_Data = Assemblies(Assembly_Name)

        Dim newInst As struct_Instance
        newInst.ClassName = Class_Name

        If (Assemblies(Assembly_Name).Error = String.Empty) Then
            newInst.instance = struct_asm.Assembly.CreateInstance(Class_Name)
            If Not newInst.instance Is Nothing Then
                newInst.Type = newInst.instance.GetType()
            Else
                newInst.Type = Nothing
                struct_asm.Error = String.Format("Class '{0}' not found.", Class_Name)
            End If

        Else
            newInst.instance = Nothing
            newInst.Type = Nothing
        End If
        struct_asm.Instances.Add(MyClass_Name, newInst)
        Assemblies(Assembly_Name) = struct_asm
    End Sub

    ''' <summary>
    ''' Call a methode
    ''' </summary>
    ''' <param name="Assembly_Name">Name of your assembly</param>
    ''' <param name="myClass_Name">Name of your Class</param>
    ''' <param name="Sub_Name">Method name</param>
    ''' <param name="Arguments">Method arguments</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Call_Function(Assembly_Name As String, myClass_Name As String, Sub_Name As String, ParamArray Arguments() As Object) As Object
        Do While IsCompiling = True
        Loop
        If (Not Assemblies.ContainsKey(Assembly_Name)) OrElse (Not Assemblies(Assembly_Name).Error = String.Empty) OrElse (Not Assemblies(Assembly_Name).Instances.ContainsKey(myClass_Name)) Then Return Nothing

        Dim myInst As struct_Instance = Assemblies(Assembly_Name).Instances(myClass_Name)

        Dim mMethode As MethodInfo = myInst.Type.GetMethod(Sub_Name)
        If mMethode Is Nothing Then
            Throw New Exception(String.Format("Method '{0}' was not found!", Sub_Name))
            Return Nothing
        End If
        Return mMethode.Invoke(myInst.instance, Arguments)
    End Function

    ''' <summary>
    ''' Try to dispose your Class
    ''' </summary>
    ''' <param name="Assembly_Name">Name of your assembly</param>
    ''' <param name="myClass_Name">Name of your Class</param>
    ''' <remarks></remarks>
    Public Sub Dispose_Class(Assembly_Name As String, myClass_Name As String)
        Do While IsCompiling = True
        Loop
        If (Not Assemblies.ContainsKey(Assembly_Name)) OrElse (Not Assemblies(Assembly_Name).Instances.ContainsKey(myClass_Name)) Then Return
        Dim myInst As struct_Instance = Assemblies(Assembly_Name).Instances(myClass_Name)

        Try
            Dim anyclass As IDisposable = DirectCast(myInst.instance, IDisposable)
            anyclass.Dispose()
            anyclass = Nothing
        Catch ex As Exception

        End Try
        myInst.instance = Nothing
        myInst.Type = Nothing
        Assemblies(Assembly_Name).Instances.Remove(myClass_Name)
    End Sub

    Private Sub OnChanged(ByVal source As Object, ByVal e As FileSystemEventArgs)
        'TODO: OnChanged has call Two time, I don't know why...
        Select Case e.ChangeType
            Case WatcherChangeTypes.Changed
                ReCompile(FileAssembly(e.FullPath))
        End Select
    End Sub

#Region "Destructor"
    Private disposedValue As Boolean

    Protected Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then

                Assemblies.Clear()
                Assemblies = Nothing

                provider.Dispose()
                provider = Nothing

                Compiler.Dispose()
                Compiler = Nothing
                params = Nothing

                FileAssembly.Clear()
                FileAssembly = Nothing
            End If

        End If
        Me.disposedValue = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
