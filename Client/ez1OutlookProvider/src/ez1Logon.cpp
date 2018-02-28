#include "..\inc\ez1OUProvider.h"
#include "..\inc\ez1Logon.h"

namespace ez1OutlookProvider
{
	ez1Logon::ez1Logon()
	{
	}


	ez1Logon::~ez1Logon()
	{
	}

	STDMETHODIMP ez1Logon::QueryInterface(REFIID riid, LPVOID * ppvObj)
	{
		*ppvObj = NULL;
		// If the interface requested is supported by this object, return a pointer
		// to the provider, with the reference count incremented by one.
		if (riid == IID_IMSLogon || riid == IID_IUnknown)
		{
			*ppvObj = (LPVOID)this;
			// Increase usage count of this object
			AddRef();
			return S_OK;
		}
		return E_NOINTERFACE;
	} // CMSProvider::QueryInterface
	STDMETHODIMP_(ULONG) ez1Logon::AddRef()
	{
		return ++m_cRef;
	}
	STDMETHODIMP_(ULONG) ez1Logon::Release()
	{
		m_cRef--;
		if (m_cRef == 0)
		{
			delete this;
		}
		return m_cRef;
	}

	////////////////////////////////////////////////////////////////////
	//	https://msdn.microsoft.com/en-us/library/cc815667.aspx
	//	Registers an object with a message store provider for notifications about changes in the message store.
	//	The message store will then send notifications about changes to the registered object.
	//	return
	//	S_OK	MAPI_E_NO_SUPPORT
	
	STDMETHODIMP ez1Logon::Advise(ULONG cbEntryID,	//[in] The size, in bytes, of the entry identifier pointed to by the lpEntryID parameter.
		LPENTRYID lpEntryID,						//[in] A pointer to the entry identifier of the object about which notifications should be generated. This object can be a folder, a message, or any other object in the message store. Alternatively, if MAPI sets the cbEntryID parameter to 0 and passes null for lpEntryID, the advise sink provides notifications about changes to the entire message store.
		ULONG ulEventMask,							//[in] An event mask of the types of notification events occurring for the object about which MAPI will generate notifications. The mask filters specific cases. Each event type has a structure associated with it that contains additional information about the event. The following table lists the possible event types along with their corresponding structures.
		LPMAPIADVISESINK lpAdviseSink,				//[in] A pointer to an advise sink object to be called when an event occurs for the session object about which notification has been requested. This advise sink object must already exist.
		ULONG_PTR FAR *lpulConnection					//[out] A pointer to a variable that upon a successful return holds the connection number for the notification registration. The connection number must be nonzero.
	) 
	{
		return S_OK;
	}

	/////////////////////////////////////////////////////////////////
	//	https://msdn.microsoft.com/en-us/library/cc765883.aspx
	// Compares two entry identifiers to determine whether they refer to the same object. MAPI refers this call to a service provider only if the unique identifiers (UIDs) in both entry identifiers to be compared are handled by that provider.
	//	return
	//	S_OK
	STDMETHODIMP ez1Logon::CompareEntryIDs(
		ULONG cbEntryID1,
		LPENTRYID lpEntryID1,
		ULONG cbEntryID2,
		LPENTRYID lpEntryID2,
		ULONG ulFlags,
		ULONG FAR * lpulResult
	)
	{
		return S_OK;
	}

	/////////////////////////////////////////////////////////////////
	//	https://msdn.microsoft.com/en-us/library/office/cc765784.aspx
	//  Returns a MAPIERROR structure that contains information about the last error that occurred for the message store object.
	//	return
	//	S_OK	MAPI_E_BAD_CHARWIDTH
	STDMETHODIMP ez1Logon::GetLastError(
		HRESULT hResult,				//[in] An HRESULT data type that contains the error value generated in the previous method call for the message store object.
		ULONG ulFlags,					// [in] A bitmask of flags that controls the type of strings returned. The following flag can be set: MAPI_UNICODE
		LPMAPIERROR FAR * lppMAPIError	//[out] A pointer to a pointer to the returned MAPIERROR structure that contains version, component, and context information for the error. The lppMAPIError parameter can be set to NULL if there is no MAPIERROR to return.
	)
	{
		return S_OK;
	}

	/////////////////////////////////////////////////////////////////
	//	https://msdn.microsoft.com/en-us/library/office/cc765740.aspx
	//	Logs off a message store provider.
	//	return
	//	S_OK
	STDMETHODIMP ez1Logon::Logoff(
		ULONG FAR * lpulFlags		//[in] Reserved; must be a pointer to zero.
	)
	{
		return S_OK;
	}

	/////////////////////////////////////////////////////////////////
	//	https://msdn.microsoft.com/en-us/library/office/cc839566.aspx
	//	Opens a folder or message object and returns a pointer to the object to provide further access.
	//	return 
	//	S_OK
	STDMETHODIMP ez1Logon::OpenEntry(
		ULONG cbEntryID,			//[in] The size, in bytes, of the entry identifier pointed to by the lpEntryID parameter.
		LPENTRYID lpEntryID,		//[in] A pointer to the address of the entry identifier of the folder or message object to open.
		LPCIID lpInterface,			//[in] A pointer to the interface identifier(IID) for the object.Passing NULL indicates that the object is cast to the standard interface for such an object.The lpInterface parameter can also be set to an identifier for an appropriate interface for the object.
		ULONG ulOpenFlags,			//[in] A bitmask of flags that controls how the object is opened. The following flags can be set:MAPI_BEST_ACCESS MAPI_DEFERRED_ERRORS MAPI_MODIFY
		ULONG FAR * lpulObjType,	//[out] A pointer to the type of the opened object.
		LPUNKNOWN FAR * lppUnk		//[out] A pointer to the pointer to the opened object.
	)
	{
		return S_OK;
	}

	/////////////////////////////////////////////////////////////////
	//	https://msdn.microsoft.com/en-us/library/office/cc842171.aspx
	//	Opens a status object.
	//	Message store providers implement the IMSLogon::OpenStatusEntry method to open a status object.
	//	This status object is then used to enable clients to call IMAPIStatus methods. 
	//	For example, clients can use the IMAPIStatus::SettingsDialog method to reconfigure the message store logon session
	//	or the IMAPIStatus::ValidateState method to validate the state of the message store logon session.
	//	return 
	//	S_OK
	STDMETHODIMP ez1Logon::OpenStatusEntry(
		LPCIID lpInterface,			//[in] A pointer to the interface identifier (IID) for the status object to open. Passing NULL indicates the standard interface for the object is returned (in this case, the IMAPIStatus interface). The lpInterface parameter can also be set to an identifier for an appropriate interface for the object.
		ULONG ulFlags,				//[in] A bitmask of flags that controls how the status object is opened. The following flag can be set:MAPI_MODIFY
		ULONG FAR * lpulObjType,	//[out] A pointer to the type of the opened object.
		LPVOID FAR * lppEntry		//[out] A pointer to the pointer to the opened object.
	)
	{
		return S_OK;
	}

	/////////////////////////////////////////////////////////////////
	//	https://msdn.microsoft.com/en-us/library/office/cc765842.aspx
	//	Removes an object's registration for notification of message store changes previously established by using a call to the IMSLogon::Advise method.
	STDMETHODIMP ez1Logon::Unadvise(
		ULONG_PTR ulConnection
	)
	{
		return S_OK;
	}
}