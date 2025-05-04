# AotNetComHost
A development-time "thunk" dll tool that enable COM support (registration, etc.) for not-yet published .NET AOT COM objects.

*But Why?* Because it's painful and not practical to develop and debug a Native AOT COM Object if you need to constantly publish it as AOT in release mode.

The following schema explains how it can be used:
![image](https://github.com/user-attachments/assets/0b4bf816-3491-4299-b2fb-12c5e78b2fff)

AotNetComHost.dll is *only useful when developping and in DEBUG mode*. Once the .NET dll is published as AOT, you shouldn't need it anymore.

To work at development time, AotNetComHost.dll and nethost.dll (provided with .NET Core files) should be placed aside the .NET dll, and AotNetComHost.dll must be renamed as the .NET dll name, followed by .something.dll, like this for example:

![image](https://github.com/user-attachments/assets/231cff72-8fc0-4ffc-a2a9-dce3a30f531b)

The *AotNetComHost* is the thunk dll project. It's pretty generic and can be used compiled as RELEASE.

The *TestComObject* is a test COM object that demonstrated how it works. Key points:
* What you should reuse are the classes in the `Hosting` folder and `IClassFactory.cs`: `ClassFactory.cs` and `ComHosting.cs`.
* `Dispatch.cs` is optional, only used with IDispatch COM objects.
* `EventProvider.cs` is a tracing tool that is optional (if you remove it, remove all its reference)
* Make sure you analyze and reproduce `TestComObject.csproj` when writing your own component as there are some subtleties in it.
* `IDispatch` support is here very limited especially around `VARIANT` types. If you need more you should consult the [DirectNAOT](https://github.com/smourier/DirectNAot) that has great a .NET AOT-compatible VARIANT (and PROPVARIANT) [wrapper class](https://github.com/smourier/DirectNAot/blob/main/DirectN.Extensions/Utilities/Variant.cs).

In DEBUG mode, you can call `regsvr32 TestComObject.comthunk.dll to register it`, `regsvr32 TestComObject.comthunk.dll /u` to unregister it.

In RELEASE mode, it's just a regular native (AOT) dll, so `regsvr32 TestComObject.dll to register it`, `regsvr32 TestComObject.dll /u` to unregister it.

There's a `test.vbs` vbscript file that demonstrates using TestComObject, in RELEASE or DEBUG, very simply (run it with `cscript.exe test.vbs`in command line):

    Set server = CreateObject("TestComObject.TestDispatchClass")
    WScript.Echo server.ComputePi() // VBS uses IDispatch interface

PS: Unlike .NET Core COM built-in support (https://github.com/dotnet/runtime/issues/45750), the thunk and ComHosting class support HKCU registration. You can use it like this `regsvr32 TestComObject.comthunk.dll /i:user /n` to register and `regsvr32 TestComObject.comthunk.dll /i:user /n /u`


