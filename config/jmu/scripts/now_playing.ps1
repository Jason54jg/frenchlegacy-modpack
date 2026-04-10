# Queries Windows Media Session for current playing track info.
# Outputs a single JSON line to stdout with title, artist, position, duration, thumbnail.
# Designed to be called repeatedly from a Java process — exits immediately after one query.

try {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime

    # Helper to await WinRT async operations
    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() |
        Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]

    Function Await($WinRtTask, $ResultType) {
        $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
        $netTask = $asTask.Invoke($null, @($WinRtTask))
        $netTask.Wait(-1) | Out-Null
        $netTask.Result
    }

    # Get session manager
    $smType = [Type]::GetType('Windows.Media.Control.GlobalSystemMediaTransportControlsSessionManager, Windows.Media.Control, ContentType=WindowsRuntime')
    $manager = Await ($smType::RequestAsync()) $smType
    $session = $manager.GetCurrentSession()

    if ($null -eq $session) {
        Write-Output '{"status":"no_session"}'
        exit 0
    }

    # Media properties
    $mpType = [Type]::GetType('Windows.Media.Control.GlobalSystemMediaTransportControlsSessionMediaProperties, Windows.Media.Control, ContentType=WindowsRuntime')
    $props = Await ($session.TryGetMediaPropertiesAsync()) $mpType

    $title = if ($props.Title) { $props.Title }  else { "" }
    $artist = if ($props.Artist) { $props.Artist } else { "" }

    # Playback info
    $pbInfo = $session.GetPlaybackInfo()

    # Timeline
    $tl = $session.GetTimelineProperties()

    # Thumbnail — use AsStream via reflection (DataReader fails on COM objects)
    $artB64 = ""
    if ($null -ne $props.Thumbnail) {
        try {
            $streamType = [Type]::GetType('Windows.Storage.Streams.IRandomAccessStreamWithContentType, Windows.Storage.Streams, ContentType=WindowsRuntime')
            $stream = Await ($props.Thumbnail.OpenReadAsync()) $streamType

            # Use .NET AsStream reflection to bridge WinRT stream to System.IO.Stream
            $asStreamMethod = [System.IO.WindowsRuntimeStreamExtensions].GetMethod(
                'AsStream',
                [Type[]]@([Windows.Storage.Streams.IRandomAccessStream])
            )

            if ($null -ne $asStreamMethod) {
                $netStream = $asStreamMethod.Invoke($null, @($stream))
                $ms = New-Object System.IO.MemoryStream
                $netStream.CopyTo($ms)
                $bytes = $ms.ToArray()
                if ($bytes.Length -gt 0 -and $bytes.Length -lt 2MB) {
                    $artB64 = [Convert]::ToBase64String($bytes)
                }
                $ms.Dispose()
                $netStream.Dispose()
            }

            $stream.Dispose()
        }
        catch { }
    }

    $result = @{
        status  = "ok"
        title   = $title
        artist  = $artist
        playing = ($pbInfo.PlaybackStatus -eq 'Playing')
        posMs   = [long]$tl.Position.TotalMilliseconds
        durMs   = [long]$tl.EndTime.TotalMilliseconds
        art     = $artB64
    }

    Write-Output (ConvertTo-Json $result -Compress)
}
catch {
    Write-Output '{"status":"error"}'
}
