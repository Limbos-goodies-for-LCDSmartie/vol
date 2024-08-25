Option Strict Off
Imports System
Imports System.Runtime.InteropServices

Module WindowsSystemAudio
    <ComImport>
    <Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IMMDeviceEnumerator
#Disable Warning IDE1006 ' Naming Styles
        Sub lpVtbl()
#Enable Warning IDE1006 ' Naming Styles
        Function GetDefaultAudioEndpoint(ByVal dataFlow As Integer,
        ByVal role As Integer, <Out> ByRef ppDevice As IMMDevice) As Integer
    End Interface

    Private NotInheritable Class MMDeviceEnumeratorFactory
        Public Shared Function CreateInstance() As IMMDeviceEnumerator
            Return CType(Activator.CreateInstance(Type.GetTypeFromCLSID(New Guid("BCDE0395-E52F-467C-8E3D-C4579291692E"))), IMMDeviceEnumerator) ' a MMDeviceEnumerator
        End Function
    End Class

    <Guid("D666063F-1587-4E43-81F1-B948E807363F"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IMMDevice
        Function Activate(
    <MarshalAs(UnmanagedType.LPStruct)> ByVal iid As Guid,
     ByVal dwClsCtx As Integer, ByVal pActivationParams As IntPtr, <Out>
    <MarshalAs(UnmanagedType.IUnknown)> ByRef ppInterface As Object) As Integer
    End Interface

    <Guid("5CDF2C82-841E-4546-9722-0CF74078229A"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IAudioEndpointVolume
        Function RegisterControlChangeNotify(ByVal pNotify As IntPtr) As Integer
        Function UnregisterControlChangeNotify(ByVal pNotify As IntPtr) As Integer
        Function GetChannelCount(ByRef pnChannelCount As UInteger) As Integer
        Function SetMasterVolumeLevel(ByVal fLevelDB As Single,
        ByVal pguidEventContext As Guid) As Integer
        Function SetMasterVolumeLevelScalar_(ByVal fLevel As Single, ByVal pguidEventContext As Guid) As Integer
        Function GetMasterVolumeLevel(ByRef pfLevelDB As Single) As Integer
        Function GetMasterVolumeLevelScalar(ByRef pfLevel As Single) As Integer
    End Interface

    Friend Sub SetVolume(ByVal Level As Integer)
        Try
            Dim deviceEnumerator As IMMDeviceEnumerator =
                                 MMDeviceEnumeratorFactory.CreateInstance()
            Dim speakers As IMMDevice = Nothing
            Dim res As Integer
            Const eRender = 0
            Const eMultimedia = 1
            deviceEnumerator.GetDefaultAudioEndpoint(eRender, eMultimedia, speakers)
            Dim Audio_EndPointVolume As IAudioEndpointVolume = Nothing
            speakers.Activate(GetType(IAudioEndpointVolume).GUID, 0, IntPtr.Zero, Audio_EndPointVolume)
            Dim ZeroGuid As New Guid()
            res = Audio_EndPointVolume.SetMasterVolumeLevelScalar_(Level / 100.0F, ZeroGuid)
        Catch ex As Exception
        End Try
    End Sub

    Friend Function GetVolume() As Integer
        Try
            Dim currentLevel As Single = 0 ' Expressed as a decimal value
            Dim deviceEnumerator As IMMDeviceEnumerator =
                                 MMDeviceEnumeratorFactory.CreateInstance()
            Dim speakers As IMMDevice = Nothing
            Dim res As Integer
            Const eRender = 0
            Const eMultimedia = 1
            deviceEnumerator.GetDefaultAudioEndpoint(eRender, eMultimedia, speakers)
            Dim Audio_EndPointVolume As IAudioEndpointVolume = Nothing
            speakers.Activate(GetType(IAudioEndpointVolume).GUID,
                              0, IntPtr.Zero, Audio_EndPointVolume)
            res = Audio_EndPointVolume.GetMasterVolumeLevelScalar(currentLevel)
            Return CInt(100 * currentLevel)  ' Returned as an Integer 0 - 100
        Catch ex As Exception
            Return -1
        End Try
    End Function
End Module

