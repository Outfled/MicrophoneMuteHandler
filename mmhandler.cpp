#include "mmhandler.h"
#include <initguid.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <functiondiscoverykeys_devpkey.h>
#include <strsafe.h>


#pragma comment ( lib, "Winmm.lib" ) 

CONST CLSID CLSID_MMDeviceEnumerator    = __uuidof(MMDeviceEnumerator);
CONST IID	IID_IMMDeviceEnumerator     = __uuidof(IMMDeviceEnumerator);
CONST IID	IID_IAudioEndpointVolume    = __uuidof(IAudioEndpointVolume);


static HRESULT GetDeviceCollection( IMMDeviceEnumerator **ppEnum, IMMDeviceCollection **ppCollection )
{
	*ppEnum			= NULL;
	*ppCollection	= NULL;

	HRESULT hResult = CoCreateInstance( CLSID_MMDeviceEnumerator, NULL, CLSCTX_ALL, IID_IMMDeviceEnumerator, (void **)ppEnum );
	if (FAILED( hResult ))
	{
		return hResult;
	}

	hResult = (*ppEnum)->EnumAudioEndpoints( eCapture, DEVICE_STATE_ACTIVE, ppCollection );
	if (FAILED( hResult ))
	{
		SafeRelease( ppEnum );
		return hResult;
	}

	return hResult;
}

DWORD
MMGetMicrophoneDeviceCount()
{
	HRESULT				hResult;
	IMMDeviceEnumerator *pEnumerator;
	IMMDeviceCollection *pCollection;
	UINT				uDevCount;

	/* Get device list */
	hResult = GetDeviceCollection( &pEnumerator, &pCollection );
	if (FAILED( hResult ))
	{
		return 0;
	}

	/* Get the list count */
	hResult = pCollection->GetCount( &uDevCount );

	/* Cleanup */
	SafeRelease( &pCollection );
	SafeRelease( &pEnumerator );

	return (DWORD)uDevCount;
}

#pragma warning ( push )
#pragma warning ( disable : 6054 ) // String might not be null-terminated
HRESULT
MMEnumerateMicrophoneDevices(
	_In_						DWORD	dwIndex,
	_Out_writes_z_( cchMax )	LPWSTR	lpszBuffer,
	_In_						SIZE_T	cchMax
)
{
	HRESULT				hResult;
	IMMDeviceEnumerator *Enumerator;
	IMMDeviceCollection *Collection;
	UINT				uMaxDevCount;
	IMMDevice			*Device;
	IPropertyStore		*PropStore;
	PROPVARIANT			vData;

	Enumerator	= NULL;
	Collection	= NULL;
	Device		= NULL;
	PropStore	= NULL;

	if (lpszBuffer == NULL || cchMax == NULL)
	{
		return E_INVALIDARG;
	}

	ZeroMemory( lpszBuffer, cchMax * sizeof(WCHAR));

	/* Get device list */
	hResult = GetDeviceCollection( &Enumerator, &Collection );
	if (FAILED( hResult ))
	{
		return hResult;
	}

	/* Get the list count */
	hResult = Collection->GetCount( &uMaxDevCount );
	if (FAILED( hResult ))
	{
		goto Cleanup;
	}

	/* Get the index device */
	hResult = Collection->Item( dwIndex, &Device );
	if (FAILED( hResult ))
	{
		goto Cleanup;
	}

	/* Open the device property story */
	hResult = Device->OpenPropertyStore( STGM_READ, &PropStore );
	if (FAILED( hResult ))
	{
		goto Cleanup;
	}

	/* Get device friendly name */
	hResult = PropStore->GetValue( PKEY_Device_FriendlyName, &vData );
	if (FAILED( hResult ))
	{
		goto Cleanup;
	}

	/* Copy name string */
	RtlCopyMemory( lpszBuffer, vData.pwszVal, (wcslen( vData.pwszVal ) + 1) * sizeof( WCHAR ) );
	
Cleanup:
	SafeRelease( &PropStore );
	SafeRelease( &Device );
	SafeRelease( &Collection );
	SafeRelease( &Enumerator );

	return hResult;
}
#pragma warning ( pop )

HRESULT
MMGetMicrophoneValue(
	_In_	LPCWSTR lpszMicrophoneName,
	_In_	int		iValueType,
	_Out_	LPVOID	lpValue
)
{
	HRESULT					hResult;
	IMMDeviceEnumerator		*Enumerator;
	IMMDeviceCollection		*Collection;
	UINT					uMaxDevCount;
	IMMDevice				*Device;
	PROPVARIANT				vData;
	IAudioEndpointVolume	*DeviceVolume;

	Enumerator	= NULL;
	Collection	= NULL;
	Device		= NULL;
	lpValue		= NULL;

	ZeroMemory( &vData, sizeof( PROPVARIANT ) );

	/* Check parameters */
	if (lpszMicrophoneName == NULL || lpValue == NULL)
	{
		return E_POINTER;
	}
	if (iValueType != MIC_DEVICE_VALUE_VOLUME && iValueType != MIC_DEVICE_VALUE_ISMUTE)
	{
		return E_INVALIDARG;
	}

	/* Get device list */
	hResult = GetDeviceCollection( &Enumerator, &Collection );
	if (FAILED( hResult ))
	{
		return 0;
	}

	/* Get the list count */
	hResult = Collection->GetCount( &uMaxDevCount );
	if (FAILED( hResult ))
	{
		goto Cleanup;
	}

	/* Enumerate all devices & find the selected device */
	for (UINT i = 0; i < uMaxDevCount; ++i)
	{
		WCHAR szName[MICROPHONE_BUF_MAX]{};

		hResult = MMEnumerateMicrophoneDevices( i, szName, ARRAYSIZE( szName ) );
		if (hResult == S_OK)
		{
			if (0 == _wcsicmp( szName, lpszMicrophoneName ))
			{
				hResult = Collection->Item( i, &Device );
				break;
			}
		}

		/* Device not found */
		if (i + i == uMaxDevCount)
		{
			hResult = E_INVALIDARG;
			break;
		}
	}

	if (FAILED( hResult )) {
		goto Cleanup;
	}

	/* Get the IAudioEndpointVolume object */
	hResult = Device->Activate( IID_IAudioEndpointVolume, CLSCTX_ALL, NULL, (LPVOID *)&DeviceVolume );
	if (FAILED( hResult ))
	{
		goto Cleanup;
	}

	/* Get the value */
	switch (iValueType)
	{
	case MIC_DEVICE_VALUE_VOLUME:
		{
			float fValue;
			WCHAR szBuf[10];

			hResult = DeviceVolume->GetChannelVolumeLevelScalar( 0, &fValue );
			if (SUCCEEDED( hResult ))
			{
				StringCbPrintf( szBuf, sizeof( szBuf ), L"%f", fValue );
				if (fValue == 1.0)
				{
					*((int *)lpValue) = 100;
				}
				else
				{
					szBuf[0] = szBuf[2];
					szBuf[1] = szBuf[3];
					szBuf[2] = L'.';
					*((int *)lpValue) = 100;
				}
			}
		}

		break;

	case MIC_DEVICE_VALUE_ISMUTE:
		hResult = DeviceVolume->GetMute( (BOOL *)lpValue );
		break;

	}

Cleanup:

	SafeRelease( &Device );
	SafeRelease( &Collection );
	SafeRelease( &Enumerator );

	return hResult;
}

HRESULT
MMSetMicrophoneValue(
	_In_ LPCWSTR	lpszMicrophoneName,
	_In_ int		iValueType,
	_In_ LPCVOID	lpValue
)
{
	HRESULT				hResult;
	IMMDeviceEnumerator *Enumerator;
	IMMDeviceCollection *Collection;
	UINT				uMaxDevCount;
	IMMDevice			*Device;
	PROPVARIANT			vData;
	IAudioEndpointVolume *DeviceVolume;

	Enumerator	= NULL;
	Collection	= NULL;
	Device		= NULL;
	ZeroMemory( &vData, sizeof( vData ) );

	/* Check parameters */
	if (lpszMicrophoneName == NULL || lpValue == NULL)
	{
		return E_POINTER;
	}
	if (iValueType != MIC_DEVICE_VALUE_VOLUME && iValueType != MIC_DEVICE_VALUE_ISMUTE)
	{
		return E_INVALIDARG;
	}

	/* Get device list */
	hResult = GetDeviceCollection( &Enumerator, &Collection );
	if (FAILED( hResult ))
	{
		return 0;
	}

	/* Get the list count */
	hResult = Collection->GetCount( &uMaxDevCount );
	if (FAILED( hResult ))
	{
		goto Cleanup;
	}

	/* Enumerate all devices & find the selected device */
	for (UINT i = 0; i < uMaxDevCount; ++i)
	{
		WCHAR szName[MICROPHONE_BUF_MAX]{};

		hResult = MMEnumerateMicrophoneDevices( i, szName, ARRAYSIZE( szName ) );
		if (hResult == S_OK)
		{
			if (0 == _wcsicmp( szName, lpszMicrophoneName ))
			{
				hResult = Collection->Item( i, &Device );
				break;
			}
		}

		/* Device not found */
		if (i + i == uMaxDevCount)
		{
			hResult = E_INVALIDARG;
			break;
		}
	}

	if (FAILED( hResult )) {
		goto Cleanup;
	}

	/* Get the IAudioEndpointVolume object */
	hResult = Device->Activate( IID_IAudioEndpointVolume, CLSCTX_ALL, NULL, (LPVOID *)&DeviceVolume );
	if (FAILED( hResult ))
	{
		goto Cleanup;
	}

	/* Get the value */
	switch (iValueType)
	{
	case MIC_DEVICE_VALUE_VOLUME:
		{
			WCHAR szBuf[10]{};
			DWORD dwVal;
			FLOAT fRealVal;

			dwVal = *(LPDWORD)lpValue;
			if (dwVal >= 100)
			{
				szBuf[0] = L'1';
				szBuf[1] = L'.';
				szBuf[2] = L'0';
			}
			else
			{
				WCHAR szAsStr[10];
				StringCbPrintf( szAsStr, sizeof( szAsStr ), L"%lu00", dwVal );

				szBuf[0] = L'0';
				szBuf[1] = L'.';
				if (dwVal < 10) {
					szBuf[2] = L'0';
					szBuf[3] = szAsStr[0];
				}
				else
				{
					szBuf[2] = szAsStr[0];
					szBuf[3] = szAsStr[1];
				}
			}

			fRealVal = (FLOAT)_wtof(szBuf);
			hResult = DeviceVolume->SetChannelVolumeLevelScalar( 0, fRealVal, NULL );
		}

		break;

	case MIC_DEVICE_VALUE_ISMUTE:
		hResult = DeviceVolume->SetMute( *(LPBOOL)lpValue, NULL );
		break;

	}

Cleanup:

	SafeRelease( &Device );
	SafeRelease( &Collection );
	SafeRelease( &Enumerator );

	return hResult;
}