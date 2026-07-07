#pragma once

#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
#include <windows.h>
#include <roapi.h>
#include <winstring.h>
#include <evntprov.h>
#include <strsafe.h>
#include <stdlib.h>

// std
#include <string>
#include <format>
#include <filesystem>

#define NETHOST_USE_AS_STATIC

#include "hostfxr.h"
#include "coreclr_delegates.h"
#include "nethost.h"
#pragma comment(lib, "nethost.lib")

// WIL (from vcpkg)
#include "wil\result.h"
#include "wil\stl.h"
#include "wil\win32_helpers.h"
#include "wil\com.h"
#include "wil\filesystem.h"

// C++/WinRT (from vcpkg)
#include "winrt\base.h"

// project globals
#include "error_codes.h" // .NET error codes
#include "wintrace.h"
