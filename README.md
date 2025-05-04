# AotNetComHost
A development-time "thunk" dll tool that enable COM support (registration, etc.) for not-yet published .NET AOT COM objects.

*But Why?* Because it's painful and not practical to develop a Native AOT COM Object if you need to constantly publish it as AOT in release mode.

The following schema explains how it can be used:
![image](https://github.com/user-attachments/assets/0b4bf816-3491-4299-b2fb-12c5e78b2fff)

AotNetComHost.dll is *only useful when developping and in DEBUG mode*. Once the .NET dll is published as AOT, you shouldn't need it anymore.

To work at development time, AotNetComHost.dll and nethost.dll (provided with .NET Core files) should be placed aside the .NET dll, renamed as the .NET dll name, followed by .something.dll, like this for example:

![image](https://github.com/user-attachments/assets/231cff72-8fc0-4ffc-a2a9-dce3a30f531b)




