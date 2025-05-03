# AotNetComHost
A tool that allows COM registration and runtime "thunk" dll for not-yet published .NET AOT COM objects.

The following schema explains how it can be used:
![image](https://github.com/user-attachments/assets/0b4bf816-3491-4299-b2fb-12c5e78b2fff)

AotNetComHost.dll is *only useful when developping and in DEBUG mode*. Once the .NET dll is published as AOT, you shouldn't need it anymore.
It should be placed aside the .NET dll, renamed as the .NET dll name, followed by .something.dll
