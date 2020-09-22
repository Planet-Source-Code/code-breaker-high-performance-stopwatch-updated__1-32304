Attribute VB_Name = "mStopWatch"
Option Explicit

'A high performance stopwatch-type module using the API.
'It's very useful to time for-loops, functions, user responses, whatever.
'It has very good accuracy, and you can call as many stopwatches as you like,
'each can be run concurrently, queried, and reset at will, and all with
'minimal overhead, regardless of the number of stopwatches you have.
'
'Available functions:
'
'StopWatchInitialize
'StopWatchStart
'StopWatchSplit
'GetStopWatchStatus
'GetStopWatchNumber
'
'Generally, to use a stopwatch:
'
'dim lHandle as long, retval as long
'
'lHandle = StopWatchInitialize   'initialize the stopwatch
'retval = StopWatchStart(lHandle)   'start
'
'Then when you want the elapsed time,
'
'retval = StopWatchSplit(lHandle)
'
'The function StopWatchSplit is so-name because the stopwatch doesn't
'actually stop, it's really more of a split time. You can keep calling the
'function for a particular stopwatch, and keep getting updated times.
'
'You can call StopWatchStart again to reset the elapsed time, and you can
'call StopWatchSplit as often as you need.
'
'by Christopher Brown, 2002, cbrown@phi.luc.edu

Private Declare Function QueryPerformanceCounter Lib "kernel32" (x As Currency) As Boolean
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (x As Currency) As Boolean
Dim Tic() As Currency, Toc() As Currency, Freq As Currency, overhead As Currency
Dim bTimerRunning() As Boolean, StopWatchCount As Long

Public Function StopWatchInitialize() As Long

'Returns a handle to a stopwatch. Pass that value to StopWatchStart to begin timing
'call this once for each timer you need

StopWatchCount = StopWatchCount + 1

ReDim Preserve Tic(StopWatchCount)
ReDim Preserve Toc(StopWatchCount)
ReDim Preserve bTimerRunning(StopWatchCount)
StopWatchInitialize = StopWatchCount

End Function

Public Function StopWatchStart(StopWatchHandle As Long) As Long

'starts a stopwatch. Returns 1 if successful, 0 if handle is invalid (stopwatch has not been initialized)

    'make sure its a valid handle
    If ((StopWatchHandle > 0) And (StopWatchHandle <= StopWatchCount)) Then
        'this is a valid handle
        StopWatchStart = 1
        bTimerRunning(StopWatchHandle) = True
        QueryPerformanceFrequency Freq
        QueryPerformanceCounter Tic(StopWatchHandle)
        QueryPerformanceCounter Toc(StopWatchHandle)
        overhead = Toc(StopWatchHandle) - Tic(StopWatchHandle)      ' determine API overhead
        QueryPerformanceCounter Tic(StopWatchHandle)        ' time loop
    Else
        StopWatchStart = 0
    End If

End Function

Public Function StopWatchSplit(StopWatchHandle As Long) As Single

'If successful, Returns a single which represents the number of seconds that have passed
'since calling StopWatchStart. I originally called this 'stopwatchstop' but that was
'misleading, since the stopwatch doesn't actually stop. It really is more of a split time,
'where you get the time, and the stopwatch keeps going.
'
'Returns 0 if StopWatchStart has not been called, -1 if handle is invalid.
    If ((StopWatchHandle > 0) And (StopWatchHandle <= StopWatchCount)) Then
        'this is a valid counter
        If bTimerRunning(StopWatchHandle) Then       'Timer has been initialized and started, proceed
            QueryPerformanceCounter Toc(StopWatchHandle)
            StopWatchSplit = (Toc(StopWatchHandle) - Tic(StopWatchHandle) - overhead) / Freq
        Else                                'Timer was initialized, but not started, return 0
            StopWatchSplit = 0
        End If
    
    Else
        StopWatchSplit = -1                  'Timer was not initialized, return -1
    End If

End Function

Public Function GetStopWatchStatus(StopWatchHandle) As Long

'Used to obtain the status of a StopWatch. Returns 1 if initialized and running,
'0 if initialized but not running (hasn't been stopwatchstart'ed), or -1 if not initialized

If ((StopWatchHandle > 0) And (StopWatchHandle <= StopWatchCount)) Then
    'valid, check if running
    If bTimerRunning(StopWatchHandle) Then
        'valid and running
        GetStopWatchStatus = 1
    Else
        GetStopWatchStatus = 0
    End If
Else
    'invalid handle, return -1
    GetStopWatchStatus = -1
End If

End Function

Public Function GetStopWatcheNumber() As Long

'If for some reason you want to know how many stopwatches you have initialized...

GetStopWatcheNumber = StopWatchCount

End Function


