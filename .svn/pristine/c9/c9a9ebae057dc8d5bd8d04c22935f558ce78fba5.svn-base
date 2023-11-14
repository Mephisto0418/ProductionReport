Imports System.CodeDom.Compiler
Imports System.Reflection

Module ReCode
    Function ReCoding(ByVal code As String)

        Dim provider As CodeDomProvider = CodeDomProvider.CreateProvider("VisualBasic")
        Dim parameters As New CompilerParameters()
        parameters.GenerateExecutable = False
        parameters.GenerateInMemory = True
        Dim results As CompilerResults = provider.CompileAssemblyFromSource(parameters, code)

        If results.Errors.HasErrors Then
            Console.WriteLine("編譯錯誤：")
            For Each [error] As CompilerError In results.Errors
                Console.WriteLine([error].ErrorText)
            Next
        Else
            Dim calculatorType As Type = results.CompiledAssembly.GetType("Calculator")
            Dim calculator As Object = Activator.CreateInstance(calculatorType)
            Dim calculateMethod As MethodInfo = calculatorType.GetMethod("Calculate")
            Dim result As Double = DirectCast(calculateMethod.Invoke(calculator, Nothing), Double)
            Console.WriteLine("計算結果： " & result)
            Return result
        End If

        Console.ReadLine()
    End Function
End Module
