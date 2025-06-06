# AotNetComHost
A development-time "thunk" dll tool that enable COM support (registration, etc.) for not-yet published .NET AOT COM in-process objects.

*But Why?* Because it's painful and not practical to develop and debug a Native AOT COM Object if you need to constantly publish it as AOT in release mode.

This [VCamNetSample](https://github.com/smourier/VCamNetSample) project demonstrates how to use the AotNetComHost binaries to show a virtual camera on Windows 11 (requires a COM object).

The following schema explains how it can be used:
![image](https://github.com/user-attachments/assets/a664fb03-ec25-4d8e-a0ca-69814f396d70)

AotNetComHost.dll is *only useful when developping and in DEBUG mode*. Once the .NET dll is published as AOT, you shouldn't need it anymore. However, provided C# classes (essentially `ComHosting` and `ClassFactory`) in the `TestComObject` project can still be used as they are fully generic and AOT-compatible.

To work at development time, AotNetComHost.dll and nethost.dll (provided with .NET Core files, located in this repo for practical use) should be placed aside the .NET dll, and AotNetComHost.dll must be renamed as the .NET dll name, followed by .something.dll, like this for example:

![image](https://github.com/user-attachments/assets/231cff72-8fc0-4ffc-a2a9-dce3a30f531b)

The *AotNetComHost* is the thunk dll project (native C++). It's pretty generic and can be used compiled as RELEASE.

The *TestComObject* is a test COM object, exposing one simple `IDispatch`-based COM interface, that demonstrates how it works. Key points:
* What you should reuse are the classes in the `Hosting` folder and `IClassFactory.cs`: `ClassFactory.cs` and `ComHosting.cs`.
* `ComHosting.cs` uses COM objects classes' `Guid` and `Progid` attributes but the exposed COM classes must be declared in `ComHosting.ComTypes` array. `ComVisible` or other .NET COM attributes are unused and irrelevant with AOT because unsupported when built-in marshaling is disabled. Note: to ensure `ProgId` will not be trimmed, your csproj needs the `BuiltInComInteropSupport` property set to `true`.
* For the native dll to find the `ComHosting` class (used as the target for COM functions), its full name must `<assemblyname>.Hosting.ComHosting` (so, located in a "Hosting" folder) or `<assemblyname>.ComHosting` (so, located in root folder).
* `Dispatch.cs` is optional, only used with IDispatch COM objects.
* `EventProvider.cs` is a tracing tool that is optional (if you remove it, remove all its reference).
* Make sure you analyze and reproduce `TestComObject.csproj` when writing your own component as there are some subtleties in it.
* `IDispatch` support is here very limited especially around `VARIANT` types. If you need more you should check the [DirectNAOT](https://github.com/smourier/DirectNAot) project that has great a .NET AOT-compatible VARIANT (and PROPVARIANT) [wrapper class](https://github.com/smourier/DirectNAot/blob/main/DirectN.Extensions/Utilities/Variant.cs).

In DEBUG mode, you can call `regsvr32 TestComObject.comthunk.dll` to register it, `regsvr32 TestComObject.comthunk.dll /u` to unregister it.

In RELEASE mode, it's just a regular native (AOT) dll, so `regsvr32 TestComObject.dll` to register it, `regsvr32 TestComObject.dll /u` to unregister it.

There's a `test.vbs` vbscript file that demonstrates using `TestComObject`, in RELEASE (so, published as native AOT) or DEBUG, very simply (run it with `cscript.exe test.vbs`in command line):

    Set server = CreateObject("TestComObject.TestDispatchClass")
    WScript.Echo server.ComputePi() // VBS uses IDispatch interface

PS: Unlike .NET Core COM built-in support (https://github.com/dotnet/runtime/issues/45750), the thunk and ComHosting class support `HKCU` registration so admin rights are not required for registration. You can use it like this `regsvr32 TestComObject.comthunk.dll /i:user /n` to register and `regsvr32 TestComObject.comthunk.dll /i:user /n /u`


