#include "..\inc\ez1OUProvider.h"
#include "..\inc\ez1MsgStore.h"

namespace ez1OutlookProvider
{
	ez1MsgStore::ez1MsgStore()
	{
	}


	ez1MsgStore::~ez1MsgStore()
	{
	}

	///////////////////////////////////////////////////////////////////////////////
	//    CMsgStore::QueryInterface()
	//
	//    Refer to the MSDN documentation for more information.
	//
	//    Purpose
	//      Returns a pointer to a interface requested if the interface is
	//      supported and implemented by this object. If it is not supported, it
	//      returns NULL
	//
	//    Return Value
	//      S_OK            If successful. With the interface pointer in *ppvObj
	//      E_NOINTERFACE   If interface requested is not supported by this object
	//
	STDMETHODIMP ez1MsgStore::QueryInterface(REFIID riid, LPVOID * ppvObj)
	{
		*ppvObj = NULL;

		// If the interface requested is supported by this object, return a pointer
		// to the provider, with the reference count incremented by one.
		if (riid == IID_IMsgStore || riid == IID_IMAPIProp || riid == IID_IUnknown)
		{
			*ppvObj = (LPVOID)this;
			// Increase usage count of this object
			AddRef();
			{
				return S_OK;
			}
		}
		return E_NOINTERFACE;
	}
	STDMETHODIMP_(ULONG) ez1MsgStore::AddRef()
	{
		return ++m_cRef;
	}

	STDMETHODIMP_(ULONG) ez1MsgStore::Release()
	{
		m_cRef--;
		if (m_cRef == 0)
		{
			delete this;
		}
		return m_cRef;
	}

	///////////////////////////////////////////////////////////////////////////////
	//    ez1MsgStore::AbortSubmit()
	//
	//    https://msdn.microsoft.com/en-us/library/office/cc815588.aspx
	//
	//    Purpose
	//      Attempts to remove a message from the outgoing queue.
	//    Return Value
	//     S_OK, MAPI_E_NOT_IN_QUEUE, MAPI_E_UNABLE_TO_ABORT always
	//
	STDMETHODIMP ez1MsgStore::AbortSubmit(ULONG	   cbEntryID,	//[in] The byte count in the entry identifier pointed to by the lpEntryID parameter.
		LPENTRYID pEntryID,										//[in] A pointer to the entry identifier of the message to remove from the outgoing queue.
		ULONG	   ulFlags)										//[in] Reserved; must be zero.
	{
		return MAPI_E_UNABLE_TO_ABORT;
	}

	///////////////////////////////////////////////////////////////////////////////
	//    ez1MsgStore::Advise()
	//
	//    https://msdn.microsoft.com/en-us/library/office/cc842238.aspx
	//
	//    Purpose
	//      Add a notification node to the list of notification subscribers for
	//      the object pointer in the entry ID passed in. In this implementation,
	//      use the MAPI notification engine to call MAPI to add a
	//      new notification subscriber to the object. The token returned from
	//      MAPI is save in the internal list of subscribers along with the
	//      notification mask. This is used when the provider's object wants to
	//      send notification about them.
	//	  Parameter
	//cbEntryID
	//	[in] The byte count in the entry identifier pointed to by the lpEntryID parameter.
	//	lpEntryID
	//	[in] A pointer to the entry identifier of the folder or message about which notifications should be generated, or null.If lpEntryID is set to NULL, Advise registers for notifications on the entire message store.
	//	ulEventMask
	//	[in] A mask of values that indicate the types of notification events that the caller is interested in and should be included in the registration.There is a corresponding NOTIFICATION structure associated with each type of event that holds information about the event.The following are valid values for the ulEventMask parameter :
	//	lpAdviseSink
	//	[in] A pointer to an advise sink object to receive the subsequent notifications.This advise sink object must have already been allocated.
	//	lpulConnection
	//	[out] A pointer to a nonzero number that represents the connection between the caller's advise sink object and the session.
	//	lpAdviseSink
	//	[in] A pointer to an advise sink object to receive the subsequent notifications.This advise sink object must have already been allocated.
	//	lpulConnection
	//	[out] A pointer to a nonzero connection number that represents the connection between the caller's advise sink object and the message store.
	//    Return Value
	//      An HRESULT	S_OK MAPI_E_NO_SUPPORT
	//	
	STDMETHODIMP ez1MsgStore::Advise(ULONG	 cbEntryID,		
		LPENTRYID		 pEntryID,							
		ULONG			 ulEventMask,						
		LPMAPIADVISESINK pAdviseSink,						
		ULONG_PTR * 	 pulConnection)
	{

		HRESULT hRes = S_OK;

		return hRes;
	}

	///////////////////////////////////////////////////////////////////////////////
	//    CMsgStore::CompareEntryIDs()
	//		https://msdn.microsoft.com/en-us/library/office/cc815456.aspx
	//
	//    Purpose
	//      This function compares the two entry ID structures passed in and if
	//      they are pointing to the same object, meaning all the members of the
	//      entry ID structure are identical, it return TRUE. Otherwise false.
	//
	//    Return Value
	//      An HRESULT. The comparison result is returned in the
	//      pulResult argument.
	//
	STDMETHODIMP ez1MsgStore::CompareEntryIDs(ULONG	   cbEntryID1,	//[in] The byte count in the entry identifier pointed to by the lpEntryID1 parameter.
		LPENTRYID pEntryID1,										//[in] A pointer to the first entry identifier to be compared.
		ULONG	   cbEntryID2,										//[in] The byte count in the entry identifier pointed to by the lpEntryID2 parameter.
		LPENTRYID pEntryID2,										//[in] A pointer to the second entry identifier to be compared.
		ULONG	   ulFlags,											//[in] Reserved; must be zero.
		ULONG *   pulResult)										//[out] A pointer to the result of the comparison. TRUE if the two entry identifiers refer to the same object; otherwise, FALSE.
	{
		HRESULT hRes = S_OK;

		return hRes;
	}

	///////////////////////////////////////////////////////////////////////////////
	//    ez1MsgStore::FinishedMsg()
	//    https://msdn.microsoft.com/en-us/library/office/cc842442.aspx
	//
	//    Purpose
	//      This message is called by the spooler to let the message store do
	//      the final processing after the message has been completely sent.
	//      During this function, the message fixes the PR_MESSAGE_FLAGS property
	//      on the submitted message and lets MAPI do final processing (i.e.
	//      Delete the message) before the message is moved out of the outgoing
	//      queue table.
	//	Parameters
	//	ulFlags
	//	[in] Reserved; must be zero.
	//	cbEntryID
	//	[in] The byte count in the entry identifier pointed to by the lpEntryID parameter.
	//	lpEntryID
	//	[in] A pointer to the entry identifier of the message to be processe
	//    Return Value
	//      An HRESULT	S_OK	MAPI_E_NO_SUPPORT
	//
	STDMETHODIMP ez1MsgStore::FinishedMsg(ULONG	ulFlags, ULONG cbEntryID, LPENTRYID pEntryID)
	{
		HRESULT hRes = S_OK;

		return hRes;
	}

	///////////////////////////////////////////////////////////////////////////////
	//    ez1MsgStore::GetOutgoingQueue()	//
	//    https://msdn.microsoft.com/en-us/library/office/cc842148.aspx
	//
	//    Purpose
	//      Return an IMAPITable interface to the caller. This table represents
	//      the outgoing queue of submitted messages. When a message is
	//      submitted, a row is added to this table with the appropiate columns.
	//		provides access to the outgoing queue table, a table that has information about all of the messages in the message store's outgoing queue. This method is called only by the MAPI spooler.
	//	  Parameters
	//	  ulFlags	[in] Reserved; must be zero.
	//	  lppTable	[out] A pointer to a pointer to the outgoing queue table.
	//    Return Value
	//      An HRESULT
	//The IMsgStore::GetOutgoingQueue method provides the MAPI spooler with access to the table that shows the message store's queue of outgoing messages. Typically, messages are placed in the outgoing queue table after their IMessage::SubmitMessage method is called. However, because the order of submission affects the order of preprocessing and submission to the transport provider, some messages that have been marked for sending might not appear in the outgoing queue table immediately.
	// by Daniel, Since we are not supporting "Submitmessage", this function doesn't neeed
	STDMETHODIMP ez1MsgStore::GetOutgoingQueue(ULONG ulFlags, LPMAPITABLE * ppTable)
	{
		HRESULT hRes = S_OK;

		return hRes;
	}

	///////////////////////////////////////////////////////////////////////////////
	//    ez1MsgStore::GetReceiveFolder()	//
	//    https://msdn.microsoft.com/en-us/library/office/cc842527.aspx
	//	Purpose
	//	Obtains the folder that was established as the destination for incoming messages of a specified message class or as the default receive folder for the message store.
	//	Parameters
	//	lpszMessageClass
	//	[in] A pointer to a message class that is associated with a receive folder.If the lpszMessageClass parameter is set to NULL or an empty string, GetReceiveFolder returns the default receive folder for the message store.
	//	ulFlags
	//	[in] A bitmask of flags that controls the type of the passed - in and returned strings.The following flag can be set :
	//		MAPI_UNICODE	The message class string is in Unicode format.If the MAPI_UNICODE flag is not set, the message class string is in ANSI format.
	//	lpcbEntryID
	//	[out] A pointer to the byte count in the entry identifier pointed to by the lppEntryID parameter.
	//	lppEntryID
	//	[out] A pointer to a pointer to the entry identifier for the requested receive folder.
	//	lppszExplicitClass
	//	[out] A pointer to a pointer to the message class that explicitly sets as its receive folder the folder pointed to by lppEntryID.This message class should either be the same as the class in the lpszMessageClass parameter, or a base class of that class.Passing NULL indicates that the folder pointed to by lppEntryID is the default receive folder for the message store.
	STDMETHODIMP ez1MsgStore::GetReceiveFolder(LPTSTR szMessageClass, ULONG ulFlags, ULONG *pcbEntryID, LPENTRYID * ppEntryID, LPTSTR *ppszExplicitClass)
	{
		HRESULT hRes = S_OK;

		return hRes;
	}

	///////////////////////////////////////////////////////////////////////////////
	//    ez1MsgStore::GetReceiveFolderTable()
	//
	//    https://msdn.microsoft.com/en-us/library/office/cc839753.aspx
	//
	//    Purpose
	//      Return the table of message class-to-folder mapping. This table is
	//      used to decide into what folder an incoming message should be placed
	//      based on its message class.
	//		Provides access to the receive folder table, a table that includes information about all of the receive folders for the message store.
	//	Parameters
	//	ulFlags
	//		[in] A bitmask of flags that controls table access.The following flags can be set :
	//		MAPI_DEFERRED_ERRORS
	//			Allows GetReceiveFolderTable to return successfully, possibly before the table is fully available to the caller.If the table is not fully available, making a subsequent table call can raise an error.
	//		MAPI_UNICODE
	//		The returned strings are in Unicode format.If the MAPI_UNICODE flag is not set, the strings are in ANSI format.
	//	lppTable
	//		[out] A pointer to a pointer to the receive folder table
	//    Return Value
	//      An HRESULT
	//
	STDMETHODIMP ez1MsgStore::GetReceiveFolderTable(ULONG ulFlags, LPMAPITABLE * ppTable)
	{
		HRESULT hRes = S_OK;

		return hRes;
	}

	///////////////////////////////////////////////////////////////////////////////
	//    ez1MsgStore::NotifyNewMail()
	//
	//    https://msdn.microsoft.com/en-us/library/office/cc842563.aspx
	//
	//    Purpose
	//      Entry point for the spooler to call into the store and tell it that a
	//      new message was placed in a folder. With this information, the
	//      spooler-side message store notifies the client-side message store,
	//      and it in turn sends a notification to the client.
	//		The IMsgStore::NotifyNewMail method is called by the MAPI spooler to inform the message store that a message is ready for delivery
	//    Return Value
	//      An HRESULT
	//
	STDMETHODIMP ez1MsgStore::NotifyNewMail(LPNOTIFICATION pNotification)
	{
		HRESULT hRes = S_OK;

		return hRes;
	}

	///////////////////////////////////////////////////////////////////////////////
	//    ez1MsgStore::OpenEntry()	//
	//    https://msdn.microsoft.com/en-us/library/office/cc815687.aspx
	//
	//    Purpose
	//      This method opens an object that exists in this message store. The object
	//      could be a folder or a message in any subfolder.
	//		The IMsgStore::OpenEntry method opens a folder or message and returns a pointer to an interface that can be used for further access.
	//    Return Value
	//      An HRESULT	S_OK	MAPI_E_NO_ACCESS	MAPI_NO_CACHE
	//
	STDMETHODIMP ez1MsgStore::OpenEntry(ULONG	   cbEntryID,		//[in] The byte count in the entry identifier pointed to by the lpEntryID parameter.
		LPENTRYID pEntryID,				//[in] A pointer to the entry identifier of the object to open, or NULL. If lpEntryID is set to NULL, OpenEntry opens the root folder for the message store.
		LPCIID pInterface,				//[in] A pointer to the interface identifier (IID) that represents the interface to be used to access the opened object. Passing NULL results in the object's standard interface (IMAPIFolder for folders and IMessage for messages) being returned.
		ULONG ulFlags,					//[in] A bitmask of flags that controls how the object is opened. The following flags can be used:	MAPI_BEST_ACCESS	MAPI_DEFERRED_ERRORS	MAPI_MODIFY
		ULONG *pulObjType,				//[out] A pointer to the type of the opened object.
		LPUNKNOWN *ppUnk)					//[out] A pointer to a pointer to the opened object.
	{
		HRESULT hRes = S_OK;

		return hRes;
	}

	///////////////////////////////////////////////////////////////////////////////
	//    ez1MsgStore::SetReceiveFolder()
	//    https://msdn.microsoft.com/en-us/library/office/cc765866.aspx
	//
	//    Purpose
	//      Associates a message class with a particular folder whose entry ID
	//      is specified by the caller. The folder is used as the "Inbox" for
	//      all messages with similar or identical message classes.
	//		Establishes a folder as the destination for incoming messages of a particular message class
	//
	//    Return Value
	//      An HRESULT
	//
	STDMETHODIMP ez1MsgStore::SetReceiveFolder(LPTSTR	szMessageClass,	//[in] A pointer to the message class that is to be associated with the new receive folder. If the lpszMessageClass parameter is set to NULL or an empty string, SetReceiveFolder sets the default receive folder for the message store.
		ULONG 	ulFlags,		//[in] A bitmask of flags that controls the type of the text in the passed-in strings. The following flag can be set:
		ULONG 	cbEntryID,		//[in] The byte count in the entry identifier pointed to by the lpEntryID parameter
		LPENTRYID pEntryID		//[in] A pointer to the entry identifier of the folder to establish as the receive folder. If the lpEntryID parameter is set to NULL, SetReceiveFolder replaces the current receive folder with the message store's default.
	)
	{
		HRESULT hRes = S_OK;

		return hRes;
	}

	///////////////////////////////////////////////////////////////////////////////
	//    ez1MsgStore::StoreLogoff()
	//    https://msdn.microsoft.com/en-us/library/office/cc815484.aspx
	//
	//    Purpose
	//      This function is used to cleanup and close up any resources owned by
	//      the store since the client plans to no longer use the message store.
	//      In this implementation, save the flags in the provider's object and return
	//      immediately. There is no special cleanup to do.
	//		Enables the orderly logoff of the message store.
	//
	//	  Parameter
	//		lpulFlags
	//		[in, out] A bitmask of flags that controls logoff from the message store.On input, all flags set for this parameter are mutually exclusive; a caller must specify only one flag per call.The following flags are valid on input :
	//		in : LOGOFF_ABORT LOGOFF_NO_WAIT	LOGOFF_ORDERLY	LOGOFF_PURGE LOGOFF_QUIET
	//		out : LOGOFF_INBOUND LOGOFF_OUTBOUND LOGOFF_OUTBOUND_QUEUE
	//    Return Value
	//      S_OK always
	//
	STDMETHODIMP ez1MsgStore::StoreLogoff(ULONG * pulFlags)
	{
		HRESULT hRes = S_OK;
		return hRes;
	}


	///////////////////////////////////////////////////////////////////////////////
	//    ez1MsgStore::Unadvise()
	//
	//    https://msdn.microsoft.com/en-us/library/office/cc765680.aspx
	//
	//    Purpose
	//      Terminates the advise link for a particular object. The connection
	//      number passed in is given to the MAPI notification engine so that it
	//      removes the connection that matches it. If the MAPI successfully
	//      removes the connection, remove it from the subscription list.
	//		Cancels the sending of notifications previously set up with a call to the IMsgStore::Advise method.
	//	  Parameters
	//		ulConnection
	//		[in] The connection number associated with an active notification registration.The value of ulConnection must have been returned by a previous call to the IMsgStore::Advise method.
	//    Return Value
	//      An HRESULT
	//
	STDMETHODIMP ez1MsgStore::Unadvise(ULONG_PTR ulConnection)
	{
		HRESULT hRes = S_OK;

		return hRes;
	}

	////IMAPIProp implement

	///////////////////////////////////////////////////////////////////////////////
	//    CMsgStore::CopyProps()
	//
	//    https://msdn.microsoft.com/en-us/library/office/cc815856.aspx.
	//
	//    Purpose
	//      Stub method. Copying the provider's properties onto
	//      another object is not supported.
	//		Copies or moves selected properties
	//	Parameter	
	//	lpIncludeProps
	//	[in] A pointer to a property tag array that specifies the properties to copy or move.PR_NULL(PidTagNull) cannot be included in the array.The lpIncludeProps parameter cannot be null.
	//	ulUIParam
	//	[in] A handle to the parent window of the progress indicator.
	//	lpProgress
	//	[in] A pointer to an implementation of a progress indicator.If null is passed in the lpProgress parameter, the progress indicator is displayed by using the MAPI implementation.The lpProgress parameter is ignored unless the MAPI_DIALOG flag is set in the ulFlags parameter.
	//	lpInterface
	//	[in] A pointer to the interface identifier(IID) that represents the interface that must be used to access the object pointed to by the lpDestObj parameter.The lpInterface parameter must not be null.
	//	lpDestObj
	//	[in] A pointer to the object to receive the copied or moved properties.
	//	ulFlags
	//	[in] A bitmask of flags that controls the copy or move operation.The following flags can be set :
	//		MAPI_DECLINE_OK
	//		If CopyProps calls the IMAPISupport::DoCopyProps method to handle the copy or move operation, it should instead return immediately with the error value MAPI_E_DECLINE_COPY.The MAPI_DECLINE_OK flag is set by MAPI to limit recursion.Clients do not set this flag.
	//		MAPI_DIALOG
	//		Displays a progress indicator.
	//		MAPI_MOVE
	//		CopyProps should perform a move operation instead of a copy operation.When this flag is not set, CopyProps performs a copy operation.
	//		MAPI_NOREPLACE
	//		Existing properties in the destination object should not be overwritten.When this flag is not set, CopyProps overwrites existing properties.
	//		lppProblems
	//	[in, out] On input, a pointer to a pointer to an SPropProblemArray structure; otherwise, null, indicating that there is no need for error information.If lppProblems is a valid pointer on input, CopyProps returns detailed information about errors in copying one or more properties.

	//    Return Value
	//      MAPI_E_NO_SUPPORT always
	//
	STDMETHODIMP ez1MsgStore::CopyProps(LPSPropTagArray		 pIncludeProps,
		ULONG_PTR			 ulUIParam,
		LPMAPIPROGRESS		 pProgress,
		LPCIID				 pInterface,
		LPVOID				 pDestObj,
		ULONG				 ulFlags,
		LPSPropProblemArray * ppProblems)
	{
		HRESULT hRes = S_OK;

		return MAPI_E_NO_SUPPORT;
	}

	///////////////////////////////////////////////////////////////////////////////
	//    CMsgStore::CopyTo()
	//
	//    Copies or moves all properties, except for specifically excluded properties.
	//	https://msdn.microsoft.com/en-us/library/office/cc839922.aspx
	//    Purpose
	//      MAPI 1.0 does not require IMsgStore to support copying itself to
	//      other store, so always return MAPI_E_NO_SUPPORT.
	//
	//    Return Value
	//      MAPI_E_NO_SUPPORT always.
	//
	STDMETHODIMP ez1MsgStore::CopyTo(ULONG ciidExclude,
		LPCIID				  rgiidExclude,
		LPSPropTagArray 	  pExcludeProps,
		ULONG_PTR			  ulUIParam,
		LPMAPIPROGRESS		  pProgress,
		LPCIID				  pInterface,
		LPVOID				  pDestObj,
		ULONG				  ulFlags,
		LPSPropProblemArray * ppProblems)
	{
		HRESULT hRes = S_OK;

		return MAPI_E_NO_SUPPORT;
	}

	///////////////////////////////////////////////////////////////////////////////
	//    CMsgStore::DeleteProps()
	//
	//    https://msdn.microsoft.com/en-us/library/office/cc839526.aspx
	//
	//    Purpose
	//      Stub method. Clients are not allowed to delete properties in the
	//      IMsgStore object.
	//
	//    Return Value
	//      MAPI_E_NO_SUPPORT always
	//
	STDMETHODIMP ez1MsgStore::DeleteProps(LPSPropTagArray pPropTagArray, LPSPropProblemArray *ppProblems)
	{
		HRESULT hRes = S_OK;

		return MAPI_E_NO_SUPPORT;
	}


	///////////////////////////////////////////////////////////////////////////////
	//    CMsgStore::GetIDsFromNames()
	//
	//   https://msdn.microsoft.com/en-us/library/office/cc839909.aspx
	//
	//    Purpose
	//      Return a property identifier for each named listed by the client. This
	//      IMsgStore implementation defers to the common IMAPIProp
	//      implementation to do the actual work.
	//		Provides the property identifiers that correspond to one or more property names.
	//
	//	Parameters
	//		cPropNames
	//		[in] The count of property names pointed to by the lppPropNames parameter.If lppPropNames is NULL, the cPropNames parameter must be 0.
	//		lppPropNames
	//		[in] A pointer to an array of property names, or NULL.Passing NULL requests property identifiers for all property names in all property sets about which the object has information.The lppPropNames parameter must not be NULL if the MAPI_CREATE flag is set in the ulFlags parameter.
	//		ulFlags
	//		[in] A bitmask of flags that indicates how the property identifiers should be returned.The following flag can be set :
	//			MAPI_CREATE
	//			Assigns a property identifier, if one has not yet been assigned, to one or more of the names included in the property name array pointed to by lppPropNames.This flag internally registers the identifier in the name - to - identifier mapping table.
	//		lppPropTags
	//		[out] A pointer to a pointer to an array of property tags that contains existing or newly assigned property identifiers.The property types for the property tags in this array are set to PT_UNSPECIFIED.
	//    Return Value
	//      An HRESULT
	//		S_OK MAPI_E_NO_SUPPORT MAPI_E_NOT_ENOUGH_MEMORY MAPI_E_TOO_BIG MAPI_W_ERRORS_RETURNED 
	//
	STDMETHODIMP ez1MsgStore::GetIDsFromNames(ULONG	   cPropNames,
		LPMAPINAMEID *ppPropNames,
		ULONG			   ulFlags,
		LPSPropTagArray *ppPropTags)
	{
		HRESULT hRes = S_OK;

		return hRes;
	}

	///////////////////////////////////////////////////////////////////////////////
	//    CMsgStore::GetNamesFromIDs()
	//
	//   https://msdn.microsoft.com/en-us/library/office/cc765793.aspx
	//
	//    Purpose
	//      Return the list of names for the identifier list supplied by the
	//      caller. This implementation defers to the common IMAPIProp
	//      implementation to do the actual work.
	//		Provides the property names that correspond to one or more property identifiers.
	//
	//	Parameters
	//		lppPropTags
	//		[in, out] On input, a pointer to an SPropTagArray structure that contains an array of property tags; otherwise, NULL, indicating that all names should be returned.The cValues member for the property tag array cannot be 0. If lppPropTags is a valid pointer on input, GetNamesFromIDs returns names for each property identifier included in the array.
	//		lpPropSetGuid
	//		[in] A pointer to a GUID, or GUID structure, that identifies a property set.The lpPropSetGuid parameter can point to a valid property set or can be NULL.
	//		ulFlags
	//		[in] A bitmask of flags that indicates the type of names to be returned.The following flags can be used(if both flags are set, no names will be returned) :
	//		MAPI_NO_IDS
	//		Requests that only names stored as Unicode strings be returned.
	//		MAPI_NO_STRINGS
	//		Requests that only names stored as numeric identifiers be returned.
	//		lpcPropNames
	//		[out] A pointer to a count of the property name pointers in the array pointed to by the lppPropNames parameter.
	//		lpppPropNames
	//		[out] A pointer to an array of pointers to MAPINAMEID structures that contains property names.
	//    Return Value
	//      An HRESULT
	//
	STDMETHODIMP ez1MsgStore::GetNamesFromIDs(LPSPropTagArray * ppPropTags,
		LPGUID 		   pPropSetGuid,
		ULONG			   ulFlags,
		ULONG *		   pcPropNames,
		LPMAPINAMEID **   pppPropNames)
	{
		HRESULT hRes = S_OK;


		return hRes;
	}
	///////////////////////////////////////////////////////////////////////////////
	//    CMsgStore::GetPropList()
	//
	//    https://msdn.microsoft.com/en-us/library/office/cc765530.aspx
	//
	//    Purpose
	//      Duplicate the tag array with ALL the available properties
	//      on this object. Properties added with SetProps() will not be returned.
	//      The must be explicitly requested in a GetProps() call.
	//
	//	Parameters
	//		ulFlags
	//		[in] A bitmask of flags that controls the format for the strings in the returned property tags.The following flag can be set :
	//		MAPI_UNICODE
	//		The returned strings are in Unicode format.If the MAPI_UNICODE flag is not set, the strings are in ANSI format.
	//		lppPropTagArray
	//		[out] A pointer to a pointer to the property tag array that contains tags for all of the object's properties.

	//    Return Value
	//      An HRESULT
	//
	STDMETHODIMP ez1MsgStore::GetPropList(ULONG ulFlags, LPSPropTagArray * ppAllTags)
	{
		HRESULT hRes = S_OK;

		return hRes;
	}
	///////////////////////////////////////////////////////////////////////////////
	//    CMsgStore::GetProps()
	//
	//   https://msdn.microsoft.com/en-us/library/office/cc765749.aspx
	//
	//	Parameters
	//		lpPropTagArray
	//		[in] A pointer to an array of property tags that identify the properties whose values are to be retrieved.The lpPropTagArray parameter must be either NULL, indicating that values for all properties of the object should be returned, or point to an SPropTagArray structure that contains one or more property tags.
	//		ulFlags
	//		[in] A bitmask of flags that indicates the format for properties that have the type PT_UNSPECIFIED.The following flag can be set :
	//			MAPI_UNICODE
	//		The string values for these properties should be returned in the Unicode format.If the MAPI_UNICODE flag is not set, the string values should be returned in the ANSI format.
	//		lpcValues
	//		[out] A pointer to a count of property values pointed to by the lppPropArray parameter.If lppPropArray is NULL, the content of the lpcValues parameter is zero.
	//		lppPropArray
	//		[out] A pointer to a pointer to the retrieved property values
	//    Purpose
	//      Return the value of the properties the client especified in the
	//      property tag array. If the tag array is NULL, the user wants all
	//      the properties.
	//
	//    Return Value
	//      An HRESULT	S_OK  MAPI_W_ERRORS_RETURNED MAPI_E_INVALID_PARAMETER 
	//
	STDMETHODIMP ez1MsgStore::GetProps(LPSPropTagArray pPropTagArray, ULONG ulFlags, ULONG *pcValues, LPSPropValue *  ppPropArray)
	{
		HRESULT hRes = S_OK;

		return hRes;
	}

	///////////////////////////////////////////////////////////////////////////////
	//    ez1MsgStore::OpenProperty()
	//
	//    https://msdn.microsoft.com/en-us/library/office/cc839910.aspx
	//
	//    Purpose
	//      Stub method. Opening properties directly on the
	//      store object is not supported.
	//		Returns a pointer to an interface that can be used to access a property.
	//
	//    Return Value
	//      MAPI_E_NO_SUPPORT always.
	//
	STDMETHODIMP ez1MsgStore::OpenProperty(ULONG 	  ulPropTag,
		LPCIID	  piid,
		ULONG 	  ulInterfaceOptions,
		ULONG 	  ulFlags,
		LPUNKNOWN * ppUnk)
	{
		HRESULT hRes = S_OK;

		return MAPI_E_NO_SUPPORT;
	}

	///////////////////////////////////////////////////////////////////////////////
	//    ez1MsgStore::SaveChanges()
	//
	//    https://msdn.microsoft.com/en-us/library/office/cc842181.aspx
	//
	//    Purpose
	//      Stub method. The store object itself is not transacted. Changes
	//      occur immediately without the need for committing changes.
	//		Makes permanent any changes that were made to an object since the last save operation. 
	//	Parameters
	//		ulFlags
	//		[in] A bitmask of flags that controls what happens to the object when the IMAPIProp::SaveChanges method is called.The following flags can be set :
	//	NON_EMS_XP_SAVE
	//		Indicates that the message has not been delivered from a Exchange Server.This flag should be used in combination with the IMAPIFolder::CreateMessage method and the ITEMPROC_FORCE flag to indicate to a PST store that the message is eligible for rules processing before the Personal Folders file(PST) store notifies any listening client that the message has arrived.This rules processing only applies to new messages that are created with IMAPIFolder::CreateMessage on a server that is not an Exchange Server, in which case the Exchange Server would have already processed rules on the message.
	//		FORCE_SAVE
	//		Changes should be written to the object, overriding any previous changes that were made to the object, and the object should be closed.Read / write permission must be set for the operation to succeed.The FORCE_SAVE flag is used after a previous call to SaveChanges returned MAPI_E_OBJECT_CHANGED.
	//		KEEP_OPEN_READONLY
	//		Changes should be committed and the object should be kept open for reading.No additional changes will be made.
	//		KEEP_OPEN_READWRITE
	//		Changes should be committed and the object should be kept open for read / write permission.This flag is usually set when the object was first opened for read / write permission.Subsequent changes to the object are allowed.
	//		MAPI_DEFERRED_ERRORS
	//		Allows SaveChanges to return successfully, possibly before the changes have been fully committed.
	//		SPAMFILTER_ONSAVE
	//		Enables spam filtering on a message that is being saved.Spam filtering support is available only if the sender’s e - mail address type is Simple Mail Transfer Protocol(SMTP), and the message is being saved to a store for a Personal Folders file(PST).
	//    Return Value
	//      S_OK always
	//
	STDMETHODIMP ez1MsgStore::SaveChanges(ULONG ulFlags)
	{		
		HRESULT hRes = S_OK;
		return hRes;
	}

	///////////////////////////////////////////////////////////////////////////////
	//    CMsgStore::SetProps()
	//
	//   https://msdn.microsoft.com/en-us/library/office/cc765899.aspx
	//
	//    Purpose
	//      Modify the value of the properties in the object. On IMsgStore
	//      objects, changes to the properties are committed immediately.
	//
	//    Return Value
	//      An HRESULT
	//
	STDMETHODIMP ez1MsgStore::SetProps(ULONG 				cValues,
		LPSPropValue			pPropArray,
		LPSPropProblemArray * ppProblems)
	{		
		HRESULT hRes = S_OK;

		return hRes;
	}

	///////////////////////////////////////////////////////////////////////////////
	//    SetLockState::SetLockState()
	//
	//    https://msdn.microsoft.com/en-us/library/office/cc765908.aspx
	//
	//    Purpose
	//      This method is called by the spooler process to lock or unlock a
	//      message for the submission process by the transports. While a
	//      message is locked, the client processes cannot access this message.
	//		Locks or unlocks a message. This method is called only by the MAPI spoole
	//
	//    Return Value
	//      An HRESULT	MSG_LOCKED	MSG_UNLOCKED
	//
	STDMETHODIMP ez1MsgStore::SetLockState(LPMESSAGE pMessageObj,
		ULONG 	ulLockState)
	{		
		HRESULT hRes = MSG_LOCKED;

		return hRes;
	}



	///////////////////////////////////////////////////////////////////////////////
	//    CMsgStore::GetLastError()
	//
	//    Refer to the MSDN documentation for more information.
	//
	//    Purpose
	//      Returns a MAPIERROR structure with an error description string about
	//      the last error that occurred in the object.
	//
	//    Return Value
	//      An HRESULT
	//
	STDMETHODIMP ez1MsgStore::GetLastError(HRESULT	   hError,
		ULONG		   ulFlags,
		LPMAPIERROR * ppMAPIError)
	{		
		return NULL;
	}





	

}
