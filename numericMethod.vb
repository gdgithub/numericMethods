'Metodos numericos para solucion de ecuaciones
'Ivan Gil   
Imports System.Math

Module Main

    Sub Main()
        'Ejercicio 1
        'MsgBox(newton(AddressOf oscilacionAmort, AddressOf oscilacionAmort_derivate, 6, 10, 0.0000001))    'b)
        'MsgBox(secante(AddressOf oscilacionAmort, 2, 1, 10, 0.00001))                                      'c)
        
        'Ejercio 2
        'MsgBox(secante(AddressOf renta, 2, 1, 100, 0.0001))

        'Ejercicio 3
        'MsgBox(biseccion(AddressOf deflexion_derivate, -500, 0, 10, 0.0000001))                            'a) Biseccion
        'MsgBox(regula_falsi(AddressOf deflexion_derivate, -500, 0, 10, 0.0000001))                         'b) Regula-Falsi
    End Sub

    Delegate Function functionMath(x As Double) As Double

    'Ejercicio 1
    Function oscilacionAmort(x As Double) As Double
        Dim k As Double = 0.7
        Dim w As Double = 4

        Return 9 * Math.Exp(-1 * k * x) * Math.Cos(w * x)
    End Function

    Function oscilacionAmort_derivate(x As Double) As Double
        Dim k As Double = 0.7
        Dim w As Double = 4

        Return -9 * Math.Exp(-1 * k * x) * (k * Math.Cos(w * x) + w * Math.Sin(w * x))
    End Function

    'Ejercicio 2
    Function renta(x as Double) as Double
        Dim P as Double = 25000
        Dim A as Double = 5500
        Dim n as Integer = 6
        
        return (P * x * (1+x)^n)/((1+x)^n - 1) - A
    End Function

    'Ejercicio 3
    Function deflexion(x as Double) as Double
        Dim w as Double = 2.5
        Dim E as Double = 50000
        Dim I as Double = 30000
        Dim L as Double = 600
        
        return (w / (120*E*I*L)) * (-1*x^5 + 2*L^2*x^3 - L^4*x) 
    End Function

    Function deflexion_derivate(x as Double) as Double
        Dim w as Double = 2.5
        Dim E as Double = 50000
        Dim I as Double = 30000
        Dim L as Double = 600
        
        return (w / (120*E*I*L)) * (-5*x^4 + 6*L^2*x^2 - L^4) 
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

    Function biseccion(f as functionMath, x0 as Double, x1 as Double, maxIter as Integer, err as Double) as Double
        Dim a,b,p,fp,fa,fb As Double
        a = x0
        b = x1

        fa = f.Invoke(a)
        fb = f.Invoke(b)
        if fa*fb > 0 Then
            Console.WriteLine("No se puede aplicar en metodo de biseccion en el intervalo dado.")
        Else
            For i as Integer = 1 To maxIter
                p = (a + b)/2
                fp = f.Invoke(p)

                If fp = 0 Or Math.Abs(b - a) < err Then
                    Console.WriteLine("Se ha alcanzado la aproximacion con la precision especificada o el cero")
                    Exit For
                Else
                    If fa*fp > 0 Then
                        a = p
                        fa = fp
                    Else
                        b = p
                        fb = fp
                    End If
                End If

            Next
        End If

        return p
    End Function

    Function regula_falsi(f as functionMath, x0 as Double, x1 as Double, maxIter as Integer, err as Double) as Double
        Dim a,b,c,fc,fa,fb As Double
        a = x0
        b = x1

        fa = f.Invoke(a)
        fb = f.Invoke(b)
        if fa*fb > 0 Then
            Console.WriteLine("No se puede aplicar en metodo de biseccion en el intervalo dado.")
        Else
            For i as Integer = 1 To maxIter
                c = (a*fb - b*fa)/(fb - fa)
                fc = f.Invoke(c)

                If fc = 0 Or Math.Abs(b - a) < err Then
                    Console.WriteLine("Se ha alcanzado la aproximacion con la precision especificada o el cero")
                    Exit For
                Else
                    If fa*fc > 0 Then
                        a = c
                        fa = fc
                    Else
                        b = c
                        fb = fc
                    End If
                End If

            Next
        End If

        return c
    End Function

End Module
