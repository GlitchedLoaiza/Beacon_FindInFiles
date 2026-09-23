Imports System.Runtime.InteropServices

Namespace Beacon
    Friend NotInheritable Class SearchConcurrencyPolicy
        Friend Const MiB As Long = 1024L * 1024
        Friend Const ReserveBytes As Long = 256 * MiB
        Friend Const WorkerBytes As Long = 64 * MiB
        Private ReadOnly _availableMemory As Func(Of Long)
        Public ReadOnly Property MaximumWorkers As Integer

        Friend Sub New(Optional forcedWorkers As Integer? = Nothing,
                       Optional processors As Func(Of Integer) = Nothing,
                       Optional availableMemory As Func(Of Long) = Nothing)
            If forcedWorkers.HasValue AndAlso forcedWorkers.Value < 1 Then Throw New ArgumentOutOfRangeException(NameOf(forcedWorkers))
            _MaximumWorkers = Math.Max(1, If(forcedWorkers, If(processors, Function() Environment.ProcessorCount).Invoke()))
            _availableMemory = If(availableMemory, AddressOf ReadAvailableMemory)
        End Sub

        Public Function AdmissionLimit() As Integer
            Dim available = AvailableBytes()
            If available <= ReserveBytes Then Return 1
            Return CInt(Math.Min(MaximumWorkers, Math.Max(1L, (available - ReserveBytes) \ WorkerBytes)))
        End Function

        Public Function CanStartHeavy(estimatedBytes As Long) As Boolean
            Dim available = AvailableBytes()
            Return available >= ReserveBytes AndAlso available - ReserveBytes >= Math.Max(WorkerBytes, estimatedBytes)
        End Function

        Private Function AvailableBytes() As Long
            Try
                Return Math.Max(0L, _availableMemory())
            Catch ex As Exception When TypeOf ex Is InvalidOperationException OrElse TypeOf ex Is System.ComponentModel.Win32Exception
                Return 0
            End Try
        End Function

        Private Shared Function ReadAvailableMemory() As Long
            Dim status As New MemoryStatus With {.Length = CUInt(Marshal.SizeOf(Of MemoryStatus)())}
            If Not GlobalMemoryStatusEx(status) Then Return 0
            Dim available = CLng(Math.Min(status.AvailablePhysical, CULng(Long.MaxValue)))
            Dim limit = GC.GetGCMemoryInfo().TotalAvailableMemoryBytes
            If limit > 0 Then available = Math.Min(available, Math.Max(0L, limit - GC.GetTotalMemory(False)))
            Return available
        End Function

        <StructLayout(LayoutKind.Sequential)>
        Private Structure MemoryStatus
            Public Length As UInteger
            Public Load As UInteger
            Public TotalPhysical As ULong
            Public AvailablePhysical As ULong
            Public TotalPageFile As ULong
            Public AvailablePageFile As ULong
            Public TotalVirtual As ULong
            Public AvailableVirtual As ULong
            Public AvailableExtendedVirtual As ULong
        End Structure

        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function GlobalMemoryStatusEx(ByRef status As MemoryStatus) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function
    End Class
End Namespace
