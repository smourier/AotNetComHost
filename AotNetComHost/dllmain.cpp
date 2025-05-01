#include "pch.h"

HMODULE _hModule = nullptr;
HRESULT _loaded = 0xFFFFFFFF;
hostfxr_initialize_for_dotnet_command_line_fn init_for_cmd_line_fptr = nullptr;
hostfxr_initialize_for_runtime_config_fn init_for_config_fptr = nullptr;
hostfxr_get_runtime_delegate_fn get_delegate_fptr = nullptr;
hostfxr_run_app_fn run_app_fptr = nullptr;
hostfxr_close_fn close_fptr = nullptr;
component_entry_point_fn _showWindowFn = nullptr;

struct registry_traits
{
	using type = HKEY;

	static void close(type value) noexcept
	{
		WINRT_VERIFY_(ERROR_SUCCESS, RegCloseKey(value));
	}

	static constexpr type invalid() noexcept
	{
		return nullptr;
	}
};

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
	auto hr = get_hostfxr_path(buffer, &buffer_size, nullptr);
	if (FAILED(hr))
		return hr;

	auto lib = LoadLibrary(buffer);
	if (!lib)
		return HRESULT_FROM_WIN32(GetLastError());

	init_for_config_fptr = (hostfxr_initialize_for_runtime_config_fn)GetProcAddress(lib, "hostfxr_initialize_for_runtime_config");
	get_delegate_fptr = (hostfxr_get_runtime_delegate_fn)GetProcAddress(lib, "hostfxr_get_runtime_delegate");
	close_fptr = (hostfxr_close_fn)GetProcAddress(lib, "hostfxr_close");

	if (!init_for_config_fptr || !get_delegate_fptr || !close_fptr)
	{
		FreeLibrary(lib);
		return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
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
	return S_OK;
}

STDAPI DllUnregisterServer()
{
	std::wstring exePath = wil::GetModuleFileNameW(_hModule).get();
	WinTrace(L"DllUnregisterServer '%s'", exePath.c_str());
	RETURN_IF_FAILED(ensure_load_hostfxr());
	return S_OK;
}

__control_entrypoint(DllExport)
STDAPI DllCanUnloadNow()
{
	WINTRACE(L"DllCanUnloadNow");
	RETURN_IF_FAILED(ensure_load_hostfxr());
	return S_OK;
}

_Check_return_
STDAPI DllGetClassObject(_In_ REFCLSID rclsid, _In_ REFIID riid, _Outptr_ LPVOID FAR* ppv)
{
	WINTRACE(L"DllGetClassObject rclsid:%s riid:%s", GUID_ToStringW(rclsid).c_str(), GUID_ToStringW(riid).c_str());
	RETURN_HR_IF_NULL(E_POINTER, ppv);
	*ppv = nullptr;
	RETURN_IF_FAILED(ensure_load_hostfxr());

	RETURN_HR(E_NOINTERFACE);
}

STDAPI DllInstall(BOOL bInstall, LPCWSTR pszCmdLine)
{
	WINTRACE(L"DllInstall bInstall:%i pszCmdLine:'%s'", bInstall, pszCmdLine);
	RETURN_IF_FAILED(ensure_load_hostfxr());
	RETURN_HR(E_NOTIMPL);
}