#pragma once

HRESULT GetTraceId(GUID* pGuid);

ULONG WinTraceRegister();
void WinTraceUnregister();

void WinTrace(PCWSTR pszFormat, ...);
void WinTrace(UCHAR level, PCWSTR string);

#ifdef _DEBUG
#define WINTRACE(...) WinTrace(__VA_ARGS__)
#else
#define WINTRACE __noop
#endif
#pragma once
