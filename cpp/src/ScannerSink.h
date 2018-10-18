#pragma once
#include "stdafx.h"

/* sink class for a Datalogic OPOS Scanner COM object
 * implements _IOPOSScannerEvents, IDispatch, IUnknown
 * OposScanner_CCO namespace is part of the OPOS.Scanner typelib
 */
class ScannerSink : public OposScanner_CCO::_IOPOSScannerEvents {

private:
	// reference counter
	LONG ref;
	// reference to scanner for events
	OposScanner_CCO::IOPOSScanner &scanner;

public:
	ScannerSink(OposScanner_CCO::IOPOSScanner &scanner) :
		scanner(scanner),
		ref(0)
	{ };

	// IUnknown methods 
	IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv);
	IFACEMETHODIMP_(ULONG) AddRef();
	IFACEMETHODIMP_(ULONG) Release();
	
	// IDispatch methods
	IFACEMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	IFACEMETHODIMP GetTypeInfo(UINT itinfo,
		LCID lcid, ITypeInfo **iti);
	IFACEMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *names,
		UINT size, LCID lcid, DISPID *rgDispId);
	IFACEMETHODIMP Invoke(DISPID dispid, REFIID riid, LCID lcid,
		WORD flags, DISPPARAMS *dispparams, VARIANT *result,
		EXCEPINFO *exceptioninfo, UINT *argerr);

private:
	// _IOPOSScannerEvents methods
	HRESULT DataEvent(
		long Status);
	HRESULT DirectIOEvent(
		long EventNumber,
		long *pData,
		BSTR *pString);
	HRESULT ErrorEvent(
		long ResultCode,
		long ResultCodeExtended,
		long ErrorLocus,
		long *pErrorResponse);
	HRESULT StatusUpdateEvent(
		long Data);
	
	/* mapping of dispids to events, queried by GetIDsOfNames() to ensure 
	 * Invoke() invokes the correct event.
	 */
	enum ScannerEvent : DISPID {
		Unused = 0,
		Data = 1,
		DirectIO = 2,
		Error = 3,
		Reserved = 4,
		StatusUpdate = 5,
	};
	std::string to_str(const ScannerEvent &e);
	std::wstring to_wstr(const ScannerEvent &e);
};