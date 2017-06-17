'Metodos numericos para solucion de ecuaciones
'Ivan Gil   
Imports System.Math

Module Main

    Sub Main()
        'Ejercicio 1
        'MsgBox(newton(AddressOf first, AddressOf first_derivate, 6, 10, 0.0000001))    'b)
        'MsgBox(secante(AddressOf first, 2, 1, 10, 0.00001))                            'c)
        
        'Ejercio 2
        'MsgBox(secante(AddressOf second, 2, 1, 100, 0.0001))
    End Sub

    Delegate Function functionMath(x As Double) As Double

    'Ejercicio 1
    Function first(x As Double) As Double
        Dim k As Double = 0.7
        Dim w As Double = 4

        Return 9 * Math.Exp(-1 * k * x) * Math.Cos(w * x)
    End Function

    Function first_derivate(x As Double) As Double
        Dim k As Double = 0.7
        Dim w As Double = 4

        Return -9 * Math.Exp(-1 * k * x) * (k * Math.Cos(w * x) + w * Math.Sin(w * x))
    End Function

    'Ejercicio 2

    Function second(x as Double) as Double
        Dim P as Double = 25000
        Dim A as Double = 5500
        Dim n as Integer = 6
        
        return (P * x * (1+x)^n)/((1+x)^n - 1) - A
    End Function

    Function newton(f As functionMath, fd As functionMath, x0 As Double, maxIter As Integer, err As Double)
        Dim a, b As Double
        a = x0
        
        For i As Integer = 1 To maxIter
            b = a - f.Invoke(a) / fd.Invoke(a)

            If Math.Abs(b - a) < err Then
                Console.WriteLine("Se ha alcanzado la aproximacion con la precision especificada")
                Exit For
            Else
                a = b
            End If
        Next 

        Return b
    End Function

    Function secante(f As functionMath, x0 As Double, x1 As Double, maxIter As Integer, err As Double) As Double
        Dim a, b, c As Double
        a = x0
        b = x1

        For i As Integer = 1 To maxIter
            c = b - f.Invoke(b) * (b - a) / (f.Invoke(b) - f.Invoke(a))
            
            If Math.Abs(c - b) < err Then
                Console.WriteLine("Se ha alcanzado la aproximacion con la precision especificada")
                Exit For
            Else
                a = b
                b = c
            End If
        Next
        
        Return c
    End Function

End Module
