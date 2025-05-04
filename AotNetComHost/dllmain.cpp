#include "pch.h"

#pragma comment(lib, "pathcch.lib")

HMODULE _hModule = nullptr;
HRESULT _loaded = 0xFFFFFFFF;
hostfxr_initialize_for_dotnet_command_line_fn init_for_cmd_line_fptr = nullptr;
hostfxr_initialize_for_runtime_config_fn init_for_config_fptr = nullptr;
hostfxr_get_runtime_delegate_fn get_delegate_fptr = nullptr;
hostfxr_run_app_fn run_app_fptr = nullptr;
hostfxr_close_fn close_fptr = nullptr;

typedef HRESULT (CORECLR_DELEGATE_CALLTYPE* dll_register_server_fn)();
typedef HRESULT (CORECLR_DELEGATE_CALLTYPE* dll_unregister_server_fn)();
typedef HRESULT (CORECLR_DELEGATE_CALLTYPE* dll_install_fn)(BOOL bInstall, LPCWSTR pszCmdLine);
typedef HRESULT (CORECLR_DELEGATE_CALLTYPE* dll_can_unload_now_fn)();
typedef HRESULT (CORECLR_DELEGATE_CALLTYPE* dll_get_class_object_fn)(_In_ REFCLSID rclsid, _In_ REFIID riid, _Outptr_ LPVOID FAR* ppv);

dll_register_server_fn dll_register_server;
dll_unregister_server_fn dll_unregister_server;
dll_install_fn dll_install;
dll_can_unload_now_fn dll_can_unload_now;
dll_get_class_object_fn dll_get_class_object;

// for communication between this and the .NET dll
typedef HRESULT(CORECLR_DELEGATE_CALLTYPE* dll_thunk_init_fn)(LPCWSTR dllPath);
dll_thunk_init_fn dll_thunk_init;

const static std::wstring GUID_ToStringW(const GUID& guid)
{
	wchar_t name[64];
	std::ignore = StringFromGUID2(guid, name, _countof(name));
	return name;
}

static HRESULT load_hostfxr()
{
	char_t buffer[2048];
	auto buffer_size = sizeof(buffer) / sizeof(char_t);
	RETURN_IF_FAILED((HRESULT)get_hostfxr_path(buffer, &buffer_size, nullptr));

	WinTrace(L"hostfxr path:'%s'", buffer);

	auto lib = LoadLibrary(buffer);
	if (!lib)
	{
		WinTrace(L"cannot load hostfxr");
		RETURN_IF_FAILED((HRESULT)CoreHostLibLoadFailure);
	}

	init_for_config_fptr = (hostfxr_initialize_for_runtime_config_fn)GetProcAddress(lib, "hostfxr_initialize_for_runtime_config");
	get_delegate_fptr = (hostfxr_get_runtime_delegate_fn)GetProcAddress(lib, "hostfxr_get_runtime_delegate");
	close_fptr = (hostfxr_close_fn)GetProcAddress(lib, "hostfxr_close");

	if (!init_for_config_fptr || !get_delegate_fptr || !close_fptr)
	{
		WinTrace(L"cannot load proc address");
		FreeLibrary(lib);
		RETURN_IF_FAILED((HRESULT)CoreHostLibMissingFailure);
	}

	auto currentDirPath = wil::GetModuleFileNameW(_hModule);
	WinTrace(L"currentDirPath: '%s'", currentDirPath.get());

	// this AotNetComHost.dll file name (after being renamed) can have any extension, it just must be named "MyNetDll.whatever.etc.whatever.dll"
	// and in this case we'll load "MyNetDll.dll" and "MyNetDll.runtimeconfig.json"
	auto fileName = wil::find_last_path_segment(currentDirPath.get());
	auto tok = wcschr(fileName, L'.');
	if (!tok)
	{
		FreeLibrary(lib);
		RETURN_IF_FAILED(HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND));
	}

	*(LPWSTR)tok = 0;
	WinTrace(L"fileName: '%s'", fileName);

	PathCchRemoveFileSpec(currentDirPath.get(), lstrlen(currentDirPath.get()));
	WinTrace(L"currentDirPath: '%s'", currentDirPath.get());

	auto rtName = std::wstring(fileName);
	rtName += L".runtimeconfig.json";

	wil::unique_cotaskmem_string rtPath;
	PathAllocCombine(currentDirPath.get(), rtName.c_str(), 0, &rtPath);

	WinTrace(L"rtPath: '%s'", rtPath.get());

	hostfxr_handle cxt = nullptr;
	auto rc = (HRESULT)init_for_config_fptr(rtPath.get(), nullptr, &cxt);
	if (FAILED(rc) || !cxt)
	{
		WinTrace(L"init_for_config_fptr failed rc:0x%08X", rc);
		close_fptr(cxt);
		FreeLibrary(lib);
		RETURN_IF_FAILED(rc);
	}

	load_assembly_fn load_assembly = nullptr;
	rc = get_delegate_fptr(cxt, hdt_load_assembly, (void**)&load_assembly);
	if (FAILED(rc) || !cxt)
	{
		WinTrace(L"get_delegate_fptr(hdt_load_assembly) failed rc:0x%08X", rc);
		close_fptr(cxt);
		FreeLibrary(lib);
		RETURN_IF_FAILED(rc);
	}

	get_function_pointer_fn get_function_pointer = nullptr;
	rc = get_delegate_fptr(cxt, hdt_get_function_pointer, (void**)&get_function_pointer);
	if (FAILED(rc) || !cxt)
	{
		WinTrace(L"get_delegate_fptr(hdt_get_function_pointer) failed rc:0x%08X", rc);
		close_fptr(cxt);
		FreeLibrary(lib);
		RETURN_IF_FAILED(rc);
	}
	close_fptr(cxt);

	auto dllName = std::wstring(fileName);
	dllName += L".dll";

	wil::unique_cotaskmem_string dllPath;
	PathAllocCombine(currentDirPath.get(), dllName.c_str(), 0, &dllPath);

	WinTrace(L"dllPath: '%s'", dllPath.get());
	RETURN_IF_FAILED((HRESULT)load_assembly(dllPath.get(), nullptr, nullptr));

	// try <assemblyname>.ComHosting, <assemblyname>
	auto typeName = std::wstring(fileName);
	typeName += L".ComHosting, ";
	typeName += fileName;

	WinTrace(L"typeName: '%s'", typeName.c_str());
	if (FAILED((HRESULT)get_function_pointer(
		typeName.c_str(),
		L"DllRegisterServer",
		UNMANAGEDCALLERSONLY_METHOD,
		nullptr,
		nullptr,
		(void**)&dll_register_server)))
	{
		// try <assemblyname>.Hosting.ComHosting, <assemblyname>
		typeName = std::wstring(fileName);
		typeName += L".Hosting.ComHosting, ";
		typeName += fileName;

		WinTrace(L"typeName: '%s'", typeName.c_str());
		RETURN_IF_FAILED_MSG((HRESULT)get_function_pointer(
			typeName.c_str(),
			L"DllRegisterServer",
			UNMANAGEDCALLERSONLY_METHOD,
			nullptr,
			nullptr,
			(void**)&dll_register_server),
			"DllRegisterServer is not exported or ComHosting type is not exposed as expected.");
	}

	RETURN_IF_FAILED_MSG((HRESULT)get_function_pointer(
		typeName.c_str(),
		L"DllUnregisterServer",
		UNMANAGEDCALLERSONLY_METHOD,
		nullptr,
		nullptr,
		(void**)&dll_unregister_server),
		"DllUnregisterServer is not exported.");

	RETURN_IF_FAILED_MSG((HRESULT)get_function_pointer(
		typeName.c_str(),
		L"DllCanUnloadNow",
		UNMANAGEDCALLERSONLY_METHOD,
		nullptr,
		nullptr,
		(void**)&dll_can_unload_now),
		"DllCanUnloadNow is not exported.");

	RETURN_IF_FAILED_MSG((HRESULT)get_function_pointer(
		typeName.c_str(),
		L"DllGetClassObject",
		UNMANAGEDCALLERSONLY_METHOD,
		nullptr,
		nullptr,
		(void**)&dll_get_class_object),
		"DllGetClassObject is not exported.");

	// DllInstall is optional
	(HRESULT)get_function_pointer(
		typeName.c_str(),
		L"DllInstall",
		UNMANAGEDCALLERSONLY_METHOD,
		nullptr,
		nullptr,
		(void**)&dll_install);

	// DllThunkInit is optional
	(HRESULT)get_function_pointer(
		typeName.c_str(),
		L"DllThunkInit",
		UNMANAGEDCALLERSONLY_METHOD,
		nullptr,
		nullptr,
		(void**)&dll_thunk_init);
	
	if (dll_thunk_init)
	{
		RETURN_IF_FAILED(dll_thunk_init(wil::GetModuleFileNameW(_hModule).get()));
	}

	return S_OK;
}

static HRESULT ensure_load_hostfxr()
{
	if (_loaded == 0xFFFFFFFF)
	{
		_loaded = load_hostfxr();
	}
	return _loaded;
}

BOOL APIENTRY DllMain(HMODULE hModule, DWORD dwReason, LPVOID lpReserved)
{
	switch (dwReason)
	{
	case DLL_PROCESS_ATTACH:
		_hModule = hModule;
		WinTraceRegister();
		WinTrace(L"DllMain DLL_PROCESS_ATTACH '%s'", GetCommandLine());
		DisableThreadLibraryCalls(hModule);

		wil::SetResultLoggingCallback([](wil::FailureInfo const& failure) noexcept
			{
				wchar_t str[2048];
				if (SUCCEEDED(wil::GetFailureLogString(str, _countof(str), failure)))
				{
					WinTrace(2, str); // 2 => error
				}
			});
		break;

	case DLL_PROCESS_DETACH:
		WinTrace(L"DllMain DLL_PROCESS_DETACH '%s'", GetCommandLine());
		WinTraceUnregister();
		break;
	}
	return TRUE;
}

STDAPI DllRegisterServer()
{
	std::wstring exePath = wil::GetModuleFileNameW(_hModule).get();
	WinTrace(L"DllRegisterServer '%s'", exePath.c_str());
	RETURN_IF_FAILED(ensure_load_hostfxr());
	RETURN_IF_FAILED(dll_register_server());
	return S_OK;
}

STDAPI DllUnregisterServer()
{
	std::wstring exePath = wil::GetModuleFileNameW(_hModule).get();
	WinTrace(L"DllUnregisterServer '%s'", exePath.c_str());
	RETURN_IF_FAILED(ensure_load_hostfxr());
	RETURN_IF_FAILED(dll_unregister_server());
	return S_OK;
}

__control_entrypoint(DllExport)
STDAPI DllCanUnloadNow()
{
	std::wstring exePath = wil::GetModuleFileNameW(_hModule).get();
	WINTRACE(L"DllCanUnloadNow '%s'", exePath.c_str());
	RETURN_IF_FAILED(ensure_load_hostfxr());
	RETURN_IF_FAILED(dll_can_unload_now());
	return S_OK;
}

_Check_return_
STDAPI DllGetClassObject(_In_ REFCLSID rclsid, _In_ REFIID riid, _Outptr_ LPVOID FAR* ppv)
{
	std::wstring exePath = wil::GetModuleFileNameW(_hModule).get();
	WINTRACE(L"DllGetClassObject '%s' rclsid:%s riid:%s", exePath.c_str(), GUID_ToStringW(rclsid).c_str(), GUID_ToStringW(riid).c_str());
	RETURN_HR_IF_NULL(E_POINTER, ppv);
	*ppv = nullptr;
	RETURN_IF_FAILED(ensure_load_hostfxr());
	RETURN_IF_FAILED(dll_get_class_object(rclsid, riid, ppv));
	return S_OK;
}

STDAPI DllInstall(BOOL bInstall, LPCWSTR pszCmdLine)
{
	std::wstring exePath = wil::GetModuleFileNameW(_hModule).get();
	WINTRACE(L"DllInstall '%s' bInstall:%i pszCmdLine:'%s'", exePath.c_str(), bInstall, pszCmdLine);
	RETURN_IF_FAILED(ensure_load_hostfxr());
	if (!dll_install)
	{
		WinTrace(L"DllInstall is not exported.");
		RETURN_IF_FAILED(E_NOTIMPL);
	}
	RETURN_IF_FAILED(dll_install(bInstall, pszCmdLine));
	return S_OK;
}