#pragma once

#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
#include <windows.h>
#include <evntprov.h>
#include <strsafe.h>
#include <stdlib.h>

// std
#include <string>
#include <format>

#define NETHOST_USE_AS_STATIC

#if _WIN64
#if _M_ARM64 
#include "runtimes\win-arm64\native\hostfxr.h"
#include "runtimes\win-arm64\native\coreclr_delegates.h"
#include "runtimes\win-arm64\native\nethost.h"
#if _DEBUG
#pragma comment(lib, "runtimes\\win-arm64\\native\\nethost.lib")
#else
#pragma comment(lib, "runtimes\\win-arm64\\native\\libnethost.lib")
#endif
#else
#include "runtimes\win-x64\native\hostfxr.h"
#include "runtimes\win-x64\native\coreclr_delegates.h"
#include "runtimes\win-x64\native\nethost.h"
#if _DEBUG
#pragma comment(lib, "runtimes\\win-x64\\native\\nethost.lib")
#else
#pragma comment(lib, "runtimes\\win-x64\\native\\libnethost.lib")
#endif
#endif // _M_ARM64
#else
#include "runtimes\win-x86\native\hostfxr.h"
#include "runtimes\win-x86\native\coreclr_delegates.h"
#include "runtimes\win-x86\native\nethost.h"
#if _DEBUG
#pragma comment(lib, "runtimes\\win-x86\\native\\nethost.lib")
#else
#pragma comment(lib, "runtimes\\win-x86\\native\\libnethost.lib")
#endif
#endif // _WIN64

// WIL (from vcpkg)
#include "wil\result.h"
#include "wil\stl.h"
#include "wil\win32_helpers.h"
#include "wil\com.h"

// C++/WinRT (from vcpkg)
#include "winrt\base.h"

// project globals
#include "wintrace.h"
