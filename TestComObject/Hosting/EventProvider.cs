﻿namespace TestComObject.Hosting;

internal sealed partial class EventProvider : IDisposable
{
    // same value as in WinTrace.cpp
    public static Guid DefaultGuid { get; set; } = new("964d4572-adb9-4f3a-8170-fcbecec27467");
    public static readonly EventProvider Default = new(DefaultGuid);

    private long _handle;
    public Guid Id { get; }

    public EventProvider(Guid id)
    {
        Id = id;
        var hr = EventRegister(id, 0, 0, out _handle);
        if (hr != 0)
            throw new Win32Exception(hr);
    }

    public void Write(string? text = null, [CallerMemberName] string? methodName = null) => EventWriteString(_handle, 0, 0, methodName + ": " + text ?? string.Empty);

    public void Dispose()
    {
        var handle = Interlocked.Exchange(ref _handle, 0);
        if (handle != 0)
        {
            _ = EventUnregister(handle);
        }
    }

    [LibraryImport("advapi32")]
    private static partial int EventRegister(in Guid ProviderId, nint EnableCallback, nint CallbackContext, out long RegHandle);

    [LibraryImport("advapi32")]
    private static partial int EventUnregister(long RegHandle);

    [LibraryImport("advapi32", StringMarshalling = StringMarshalling.Utf16)]
    private static partial int EventWriteString(long RegHandle, byte Level, long Keyword, string String);
}
