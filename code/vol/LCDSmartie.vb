'add rererence to coreaudioapi.dll
Imports CoreAudioApi
Imports System.Runtime.InteropServices
Imports System.Text
Imports Microsoft.VisualBasic.Devices
Imports System.Media
Imports System.Runtime.InteropServices.WindowsRuntime
Public Class LCDSmartie

    Dim prevVolLevel As Integer

    Private Function getmute() As Boolean
        Dim devenum As New MMDeviceEnumerator
        Dim device As MMDevice = devenum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia)
        If device.AudioEndpointVolume.Mute = True Then
            Return True
        Else
            Return False
        End If
    End Function


    Dim pluginName As String = "vol"


    Public Function function1(ByVal param1 As String, ByVal param2 As String) As String
        If LCase(param1) = "about" Then
            Return "returns volume level"

        Else
            Dim currentVolume As Integer = GetVolume()
            If getmute() = True Then 'if the volume is muted
                If param2 = "" Then 'check if used passed parameter 2 to handle mute if not 
                    Return currentVolume
                Else 'when param2 value is passed then return this instead of the volume on mute 
                    Return param2
                End If
            Else

                Return currentVolume
            End If

        End If

    End Function


    Public Function function2(ByVal param1 As String, ByVal param2 As String) As String
        If LCase(param1) = "about" Then
            Return "returns volume level in two digit"
        Else

            Dim currentVolume As Integer = GetVolume()
            If getmute() = True Then 'if the volume is muted
                If param2 = "" Then 'check if used passed parameter 2 to handle mute if not 
                    Return currentVolume.ToString("00")
                Else 'when param2 value is passed then return this instead of the volume on mute 
                    Return param2
                End If
            Else

                Return currentVolume.ToString("00")
            End If

        End If

        '    Dim currentVolume As Integer = GetVolume()

        '    Return currentVolume.ToString("00")
        'End If


    End Function

    Public Function function3(ByVal param1 As String, ByVal param2 As String) As String
        If LCase(param1) = "about" Then
            Return "returns volume level in three digit"

        Else

            Dim currentVolume As Integer = GetVolume()
            If getmute() = True Then 'if the volume is muted
                If param2 = "" Then 'check if used passed parameter 2 to handle mute if not 
                    Return currentVolume.ToString("000")
                Else 'when param2 value is passed then return this instead of the volume on mute 
                    Return param2
                End If
            Else

                Return currentVolume.ToString("000")
            End If

        End If
        '    Dim currentVolume As Integer = GetVolume()

        '    Return currentVolume.ToString("000")

        'End If

    End Function


    Public Function function4(ByVal param1 As String, ByVal param2 As String) As String


        If LCase(param1) = "about" Then
            Return "returns volume with percent symbol"
        Else

            Dim currentVolume As Integer = GetVolume()
            If getmute() = True Then 'if the volume is muted
                If param2 = "" Then 'check if used passed parameter 2 to handle mute if not 
                    Return currentVolume & "%"
                Else 'when param2 value is passed then return this instead of the volume on mute 
                    Return param2
                End If
            Else

                Return currentVolume & "%"
            End If

        End If
        '    Dim currentVolume As Integer = GetVolume()
        '    Return currentVolume & "%"
        'End If


    End Function


    Public Function function5(ByVal param1 As String, ByVal param2 As String) As String


        If LCase(param1) = "about" Then
            Return "returns mute status"
        Else

            If getmute() Then
                If param2 = "" Then
                    Return "muted"
                Else
                    Return 0
                End If

            Else
                If param2 = "" Then
                    Return "unmuted"
                Else
                    Return 1
                End If
            End If
        End If


    End Function


    Public Function function19(ByVal param1 As String, ByVal param2 As String) As String
        If LCase(param1) = "about" Then
            Return "Returns 1 when a volume change is detected"

        Else
            Dim currentVolume As Integer = GetVolume()
            If currentVolume <> prevVolLevel Then
                Return 1
                prevVolLevel = currentVolume
            Else
                Return 0
            End If

        End If


    End Function

    Public Function function20(ByVal param1 As String, ByVal param2 As String) As String
        If LCase(param1) = "about" Then
            Return "Created by Nikos Georgousis - 2024"

        Else
            Return pluginName & " v1.3 by Limbo"
        End If


    End Function



    Public Function GetMinRefreshInterval() As Integer

        Return 100 '100 ms (around 10 times a second)

    End Function


    Public Function SmartieDemo()

        Dim demolist As New StringBuilder()
        demolist.AppendLine("-- " & pluginName & " plugin for LCD Smartie")
        demolist.AppendLine("returns system volume level")
        demolist.AppendLine("This plugin can provide up to 10 values per second.")
        demolist.AppendLine("------ Function1 ------")
        demolist.AppendLine(">  returns volume level   <")
        demolist.AppendLine(" Current value  $dll(vol,1,,)")
        demolist.AppendLine("------ Function2 ------")
        demolist.AppendLine(">  returns volume level in two digit <")
        demolist.AppendLine("Current value  $dll(vol,2,,)")
        demolist.AppendLine("------ Function3 ------")
        demolist.AppendLine(">   returns volume level in three digit   <")
        demolist.AppendLine("Current value  $dll(vol,3,,)")
        demolist.AppendLine("------ Function4 ------")
        demolist.AppendLine(">  returns volume level followed by percent symbol   <")
        demolist.AppendLine("Current value  $dll(vol,4,,)")

        demolist.AppendLine("------ Function5 ------")
        demolist.AppendLine(">  returns volume level followed by percent symbol   <")
        demolist.AppendLine("Get text value  $dll(vol,5,,)")
        demolist.AppendLine("Get numerical value 0=muted 1=unmuted  $dll(vol,5,,any_value)")

        demolist.AppendLine("------ Function19 ------")
        demolist.AppendLine(">  returns 1 when a volume change is detected   <")
        demolist.AppendLine("Get text value  $dll(vol,19,,)")
        demolist.AppendLine("This is useful to trigger actions and switch screens only when volume is changed")

        demolist.AppendLine("------ Function20 ------")
        demolist.AppendLine(">  About  <")
        demolist.AppendLine("Get about info:  $dll(vol,20, , )")
        demolist.AppendLine("Note: If param2 is passed on functions 1~4 then this value will be returned when the volume is muted")
        demolist.AppendLine("Example: $dll(vol,2,,muted)")
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
        Return "Developer: Nikos Georgousis (limbo)" & vbNewLine & "Version: 1.3 "
    End Function



End Class
