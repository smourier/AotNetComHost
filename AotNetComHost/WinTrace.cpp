#include "pch.h"

// we don't use OutputDebugString because it's 100% crap, truncating, slow, etc.
// use WpfTraceSpy https://github.com/smourier/TraceSpy to see these traces (configure an ETW Provider with guid set to 964d4572-adb9-4f3a-8170-fcbecec27467)
static GUID GUID_WinTraceProvider = { 0x964d4572,0xadb9,0x4f3a,{0x81,0x70,0xfc,0xbe,0xce,0xc2,0x74,0x67} };

REGHANDLE _traceHandle = 0;

HRESULT GetTraceId(GUID* pGuid)
{
	if (!pGuid)
		return E_INVALIDARG;

	*pGuid = GUID_WinTraceProvider;
	return S_OK;
}

ULONG WinTraceRegister()
{
	return EventRegister(&GUID_WinTraceProvider, nullptr, nullptr, &_traceHandle);
}

void WinTraceUnregister()
{
	auto h = _traceHandle;
	if (h)
	{
		_traceHandle = 0;
		EventUnregister(h);
	}
}

void WinTrace(PCWSTR format, ...)
{
	if (!_traceHandle)
		return;

	WCHAR szTrace[2048];
	va_list args;
	va_start(args, format);
	// add '00000000:' before all traces
	StringCchPrintf(szTrace, (size_t)(9 + 1), L"%08X:", GetCurrentThreadId());
	StringCchVPrintfW(((LPWSTR)szTrace) + 9, _countof(szTrace) - 10, format, args);
	va_end(args);
	EventWriteString(_traceHandle, 0, 0, szTrace);
}

void WinTrace(UCHAR level, PCWSTR string)
{
	if (!_traceHandle)
		return;

	EventWriteString(_traceHandle, level, 0, string);
}