
#include "..\inc\ez1OUProvider.h"
#include "..\inc\ez1MSProvider.h"

namespace ez1OutlookProvider
{
	ez1MSProvider::ez1MSProvider(ULONG ulFlags)
	{
		Log(true, "ez1MSProvider::QCSProvider\n");
		m_ulFlags = ulFlags;
	}


	ez1MSProvider::~ez1MSProvider()
	{
	}

	HRESULT ez1MSProvider::Create(ULONG ulFlags, ez1MSProvider **lppMSProvider)
	{
		Log(true, "ez1MSProvider::Create\n");
		ez1MSProvider *lpMSProvider = new ez1MSProvider(ulFlags);
		return lpMSProvider->QueryInterface(IID_IMSProvider, (void **)lppMSProvider);
	}

	///////////////////////////////////////////////////////////////////////////////
	//    CMSProvider::QueryInterface()
	//
	//    Refer to the MSDN documentation for more information.
	//
	//    Purpose
	//      Returns a pointer to a interface requested if the interface is
	//      supported and implemented by this object. If it is not supported, it
	//      returns NULL.
	//
	//    Return Value
	//      S_OK            If successful. With the interface pointer in *ppvObj
	//      E_NOINTERFACE   If interface requested is not supported by this object
	//
	STDMETHODIMP ez1MSProvider::QueryInterface(REFIID refiid, void **lppInterface)
	{
		Log(true, "ez1MSProvider::QueryInterface\n");
		*lppInterface = NULL;

		// If the interface requested is supported by this object, return a pointer
		// to the provider, with the reference count incremented by one.
		REGISTER_INTERFACE(IID_IMSProvider, this);
		REGISTER_INTERFACE(IID_IUnknown, this);
		
		return E_NOINTERFACE;
	} // CMSProvider::QueryInterface
	STDMETHODIMP_(ULONG) ez1MSProvider::AddRef()
	{
		Log(true, "ez1MSProvider::AddRef\n");
		return ++m_cRef;
	}
	STDMETHODIMP_(ULONG) ez1MSProvider::Release()
	{
		Log(true, "ez1MSProvider::Release\n");
		m_cRef--;
		if (m_cRef == 0)
		{
			delete this;
		}
		return m_cRef;
	}
	

	///////////////////////////////////////////////////////////////////////////////
	//		ez1MSProvider::CompareStoreIDs()
	//		https://msdn.microsoft.com/en-us/library/office/cc842454.aspx
	//    Refer to the MSDN documentation for more information.
	//
	//    Purpose
	//		Compares two message store entry identifiers to determine whether they refer to the same store object.
	//      Give two entry IDs, compare them to find out if they are the
	//      same. (i.e. they refer to the same object)
	//
	//		cbEntryID1  
	//		[in] The size, in bytes, of the entry identifier pointed to by the lpEntryID1 parameter.
	//		lpEntryID1  
	//		[in] A pointer to the first entry identifier to be compared.
	//		cbEntryID2  
	//		[in] The size, in bytes, of the entry identifier pointed to by the lpEntryID2 parameter.
	//		lpEntryID2  
	//		[in] A pointer to the second entry identifier to be compared.
	//		ulFlags  
	//		[in] Reserved; must be zero.
	//		lpulResult  
	//		[out] A pointer to the returned result of the comparison. TRUE if the two entry identifiers refer to the same object; otherwise, FALSE.
	//    Return Value
	//      An HRESULT
	//
	STDMETHODIMP ez1MSProvider::CompareStoreIDs(ULONG	cbEntryID1, LPENTRYID pEntryID1, ULONG cbEntryID2, LPENTRYID pEntryID2, ULONG ulFlags, ULONG *pulResult)
	{
		Log(true, "ez1MSProvider::CompareStoreIDs\n");
		HRESULT hRes = S_OK;

		return hRes;
	}



	///////////////////////////////////////////////////////////////////////////////
	//    ez1MSProvider::Logon()
	//
	//    Refer to the MSDN documentation for more information.
	//		https://msdn.microsoft.com/en-us/library/office/cc842201.aspx
	//    Purpose
	//      This method logs on the store by opening the message store database
	//      and allocating the the IMsgStore and IMSLogon objects that are
	//      returned to MAPI. During this function the provider gets the
	//      configuration properties stored in its corresponding profile section.
	//      This infomation include the location of the database file and other
	//      options important at logon time.
	//      If the client is trying to logon onto a message store database file
	//      which has already been open then don't create a new object, but
	//      simply return an AddRef()'ed pointer to the IMsgStore logon object
	//      and a NULL IMSLogon object. These are the rules of MAPI.
	//      If new objects are created they are added to the chain of open
	//      message store objects in this IMSProvider.
	//
	//    Return Value
	//      An HRESULT
	//		S_OK  MAPI_E_FAILONEPROVIDER MAPI_E_LOGON_FAILED MAPI_E_UNCONFIGURED MAPI_E_USER_CANCEL MAPI_E_UNKNOWN_CPID MAPI_E_UNKNOWN_LCID MAPI_W_ERRORS_RETURNED 
	STDMETHODIMP ez1MSProvider::Logon(
		LPMAPISUP	  pSupObj,				//[in] A pointer to the current MAPI support object for the message store.
		ULONG_PTR	  ulUIParam,			//[in] A handle to the parent window of any dialog boxes or windows this method displays. 
		LPTSTR		  pszProfileName,		//[in] A pointer to a string that contains the name of the profile being used for store provider logon. This string can be displayed in dialog boxes, written out to a log file, or simply ignored. It must be in Unicode format if the MAPI_UNICODE flag is set in the ulFlags parameter.
		ULONG		  cbEntryID,			//[in] The size, in bytes, of the entry identifier pointed to by the lpEntryID parameter.
		LPENTRYID	  pEntryID,				//[in] A pointer to the entry identifier for the message store. Passing null in lpEntryID indicates that a message store has not yet been selected and that dialog boxes that enable the user to select a message store can be presented.
		ULONG		  ulFlags,				//[in] A bitmask of flags that controls how the logon is performed. The following flags can be set:  MAPI_DEFERRED_ERRORS MAPI_UNICODE  MDB_NO_DIALOG MDB_NO_MAIL MDB_TEMPORARY MDB_WRITE 
		LPCIID		  pInterface,			//[in] A pointer to the interface identifier (IID) for the message store to log on to. Passing null indicates the MAPI interface for the message store (IMsgStore) is returned. The lpInterface parameter can also be set to an identifier for an appropriate interface for the message store (for example, IID_IUnknown or IID_IMAPIProp).
		ULONG * 	  pcbSpoolSecurity,		//[out] A pointer to the variable in which the store provider returns the size, in bytes, of the validation data in the lppbSpoolSecurity parameter.
		LPBYTE *	  ppbSpoolSecurity,		//[out] A pointer to the pointer to the returned validation data. This validation data is provided so the IMSProvider::SpoolerLogon method can log the MAPI spooler on to the same store as the message store provider.
		LPMAPIERROR * ppMAPIError,			//[out] A pointer to a pointer to the returned MAPIERROR structure, if any, that contains version, component, and context information for an error.The lppMAPIError parameter can be set to null if there is no MAPIERROR structure to return.
		LPMSLOGON *   ppMSLogon,			//[out] A pointer to the pointer to the message store logon object for MAPI to log on to.
		LPMDB * 	  ppMDB)				//[out] A pointer to the pointer to the message store object for the MAPI spooler and client applications to log on to.
	{
		Log(true, "ez1MSProvider::Logon\n");
		//HRESULT hr = S_OK;
		//LPMDB lpPSTMDB = NULL;
		//sGlobalProfileProps	sProfileProps;


		//try
		//{
		//	// Always suppress UI when running in a service
		//	if (m_ulFlags & MAPI_NT_SERVICE)
		//		ulFlags |= MDB_NO_DIALOG;

		//	// If the EntryID is not configured, return MAPI_E_UNCONFIGURED, this will
		//	// cause MAPI to call our configuration entry point (MSGServiceEntry)
		//	if (pEntryID == NULL) {
		//		hr = MAPI_E_UNCONFIGURED;
		//		goto exit;
		//	}
		//	if (pcbSpoolSecurity)
		//		*pcbSpoolSecurity = 0;
		//	if (ppbSpoolSecurity)
		//		*ppbSpoolSecurity = NULL;


		//	hr = QCSMapiUtil::GetGlobalProfileProperties(pSupObj, &sProfileProps);
		//	if (hr != hrSuccess)
		//		goto exit;

		//	exit:
		//}
		//catch (const std::exception& e)
		//{

		//}
		

		// need to create LPMSLOGON object
		return MAPI_E_LOGON_FAILED;
	} // ez1MSProvider::Logon

	///////////////////////////////////////////////////////////////////////////////
	//    ez1MSProvider::Shutdown()
	//
	//    Refer to the MSDN documentation for more information.
	//	https://msdn.microsoft.com/en-us/library/office/cc815596.aspx
	//
	//    Purpose
	//      To terminate all session objects attached to this provider object
	// Parameter
	//	pulFlags :[in] Reserved; must be a pointer to zero
	//    Return Value
	//      S_OK always.
	//
	STDMETHODIMP ez1MSProvider::Shutdown(ULONG * pulFlags)
	{
		Log(true, "ez1MSProvider::Shutdown\n");
		HRESULT hRes = S_OK;

		return hRes;
	} // CMSP


	///////////////////////////////////////////////////////////////
	//	
	//	https://msdn.microsoft.com/en-us/library/office/cc842063.aspx
	//	Logs the MAPI spooler on to a message store.
	//	The MAPI spooler calls the IMSProvider::SpoolerLogon method to log on to a message store
	//	The MAPI spooler should use the message store object returned by the message store provider in the lppMDB parameter during and after logon

	STDMETHODIMP ez1MSProvider::SpoolerLogon(LPMAPISUP	  pSupObj,	//[in] A pointer to the MAPI support object for the message store.
		ULONG_PTR	  ulUIParam,									//[in] A handle to the parent window of any dialog boxes or windows this method displays. 
		LPTSTR		  pszProfileName,								//[in] A pointer to a string that contains the name of the profile being used for the MAPI spooler logon. This string can be displayed in dialog boxes, written out to a log file, or simply ignored. It must be in Unicode format if the MAPI_UNICODE flag is set in the ulFlags parameter.
		ULONG		  cbEntryID,									//[in] The size, in bytes, of the entry identifier pointed to by the lpEntryID parameter.
		LPENTRYID	  lpEntryID,										//[in] A pointer to the entry identifier for the message store.Passing NULL in the lpEntryID parameter indicates that a message store has not yet been selected and that dialog boxes that enable the user to select a message store can be presented.
		ULONG		  ulFlags,										//[in] A bitmask of flags that controls how the logon is performed. The following flags can be set: MAPI_DEFERRED_ERRORS MAPI_UNICODE MDB_NO_DIALOG MDB_WRITE 
		LPCIID		  pInterface,									//[in] A pointer to the interface identifier (IID) for the message store to log on to. Passing NULL indicates the MAPI interface for the message store (IMsgStore) is returned. The lpInterface parameter can also be set to an identifier for an appropriate interface for the message store (for example IID_IUnknown or IID_IMAPIProp).
		ULONG		  cbSpoolSecurity,								//[in] A pointer to the size, in bytes, of validation data in the lppbSpoolSecurity parameter.
		LPBYTE		  pbSpoolSecurity,								//[in] A pointer to a pointer to validation data. The SpoolerLogon method uses this data to log the MAPI spooler on to the same store as the message store provider previously logged on to by using the IMSProvider::Logon method.
		LPMAPIERROR * ppMAPIError,									//[out] A pointer to a pointer to the returned MAPIERROR structure, if any, that contains version, component, and context information for an error. The lppMAPIError parameter can be set to NULL if there is no MAPIERROR structure to return.
		LPMSLOGON *   ppMSLogon,									//[out] A pointer to the pointer to the message store logon object for MAPI to log on to.
		LPMDB * 	  ppMDB)										//[out] A pointer to the pointer to the message store object for the MAPI spooler and client applications to log on to	
	{
		Log(true, "ez1MSProvider::SpoolerLogon\n");
		HRESULT hr = S_OK;

				
		return hr;
	}
		 	
}