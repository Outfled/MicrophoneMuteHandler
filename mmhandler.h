//------------------------------------------------
//
// WindowsMicrophoneMute.h
// Author: Outfled
//
// Header file for mmhandler.dll
// 
//------------------------------------------------

#ifndef MM_HANDLER_H
#define MM_HANDLER_H

#include <Windows.h>

#define MICMUTEAPI			__stdcall
#define MICMUTEEXPORT		__declspec(dllexport)

template<class T> VOID SafeRelease( T **ppObj )
{
	if (ppObj && *ppObj)
	{
		(*ppObj)->Release();
		*ppObj = nullptr;
	}
}

#ifdef __cplusplus
extern "C" {
#endif //cplusplus


#define MICROPHONE_BUF_MAX			256

#define MIC_DEVICE_VALUE_VOLUME		1
#define MIC_DEVICE_VALUE_ISMUTE		2


	//----------------------------------------
	// Get Number of Connected Microphone Devices
	MICMUTEEXPORT
		DWORD
		MICMUTEAPI
		MMGetMicrophoneDeviceCount();

	//----------------------------------------
	// Get Microphone Device Name by Index
	MICMUTEEXPORT
		HRESULT
		MICMUTEAPI
		MMEnumerateMicrophoneDevices(
			_In_						DWORD	dwIndex,
			_Out_writes_z_( cchMax )	LPWSTR	lpszBuffer,
			_In_						SIZE_T	cchMax
		);

	//----------------------------------------
	// Get Microphone Value
	MICMUTEEXPORT
		HRESULT
		MICMUTEAPI
		MMGetMicrophoneValue(
			_In_	LPCWSTR lpszMicrophoneName,
			_In_	int		iValueType,
			_Out_	LPVOID	lpValue
		);

	//----------------------------------------
	// Set Microphone Value
	MICMUTEEXPORT
		HRESULT
		MICMUTEAPI
		MMSetMicrophoneValue(
			_In_ LPCWSTR	lpszMicrophoneName,
			_In_ int		iValueType,
			_In_ LPCVOID	lpValue
		);

#ifdef __cplusplus
}
#endif // __cplusplus


#endif // !MM_HANDLER_H
