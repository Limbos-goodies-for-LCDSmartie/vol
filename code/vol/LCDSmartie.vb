Imports System.Text

Public Class LCDSmartie


    Dim pluginName As String = "vol"


    Public Function function1(ByVal param1 As String, ByVal param2 As String) As String
        If LCase(param1) = "about" Then
            Return "returns volume level"

        Else
            Dim currentVolume As Integer = GetVolume()
            Return currentVolume
        End If

    End Function


    Public Function function2(ByVal param1 As String, ByVal param2 As String) As String
        If LCase(param1) = "about" Then
            Return "returns volume level in two digit"
        Else
            Dim currentVolume As Integer = GetVolume()

            Return currentVolume.ToString("00")
        End If


    End Function

    Public Function function3(ByVal param1 As String, ByVal param2 As String) As String
        If LCase(param1) = "about" Then
            Return "returns volume level in three digit"

        Else
            Dim currentVolume As Integer = GetVolume()

            Return currentVolume.ToString("000")

        End If

    End Function


    Public Function function4(ByVal param1 As String, ByVal param2 As String) As String


        If LCase(param1) = "about" Then
            Return "returns volume with percent symbol"
        Else
            Dim currentVolume As Integer = GetVolume()
            Return currentVolume & "%"
        End If


    End Function





    Public Function function20(ByVal param1 As String, ByVal param2 As String) As String
        If LCase(param1) = "about" Then
            Return "Created by Nikos Georgousis - 2024"

        Else

            Return pluginName & " v1.1 by Limbo"
        End If


    End Function



    Public Function GetMinRefreshInterval() As Integer

        Return 200 ' 200 ms (around 5 times a second)

    End Function


    Public Function SmartieDemo()



        Dim demolist As New StringBuilder()
        demolist.AppendLine("-- " & pluginName & " plugin for LCD Smartie")
        demolist.AppendLine("returns system volume level")
        demolist.AppendLine("This plugin can provide up to 5 values per second.")
        demolist.AppendLine("------ Function1 ------")
        demolist.AppendLine(">  returns volume level   <")
        demolist.AppendLine("Current value  $dll(vol,1, , )")
        demolist.AppendLine("------ Function2 ------")
        demolist.AppendLine(">  returns volume level in two digit <")
        demolist.AppendLine("Current value  $dll(vol,2, , )")
        demolist.AppendLine("------ Function3 ------")
        demolist.AppendLine(">   returns volume level in three digit   <")
        demolist.AppendLine("Current value  $dll(vol,1, , )")
        demolist.AppendLine("------ Function4 ------")
        demolist.AppendLine(">  returns volume level followed by percent symbol   <")
        demolist.AppendLine("Current value  $dll(vol,2, , )")
        demolist.AppendLine("------ Function20 ------")
        demolist.AppendLine(">  About <")
        demolist.AppendLine("Get about info:  $dll(vol,20, , )")
        demolist.AppendLine("")
        demolist.AppendLine("** Visit **")
        demolist.AppendLine("> Forums")
        demolist.AppendLine("https://www.lcdsmartie.org")
        demolist.AppendLine("> Active development branch (latest version)")
        demolist.AppendLine("https://github.com/LCD-Smartie/LCDSmartie")
        demolist.AppendLine("> My plugins can be found on:")
        demolist.AppendLine("https://github.com/Limbos-goodies-for-LCDSmartie")
        demolist.AppendLine("")
        '  demolist.AppendLine("Note: This is a note for " & pluginName)


        Dim result As String = demolist.ToString()
        Return result
    End Function

    Public Function SmartieInfo()
        Return "Developer: Nikos Georgousis (limbo)" & vbNewLine & "Version: 1.1 "
    End Function



End Class
