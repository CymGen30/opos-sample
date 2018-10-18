#include "stdafx.h"
#include "scannersink.h"

// IUnknown methods
/* Returns a pointer to a scanner sink interface for this object, allowing
 * the COM object to call this.Invoke()
 */
IFACEMETHODIMP 
ScannerSink::QueryInterface(REFIID riid, void **ppv) {
	*ppv = nullptr;
	IID id = __uuidof(OposScanner_CCO::_IOPOSScannerEvents);
	HRESULT hr = E_NOINTERFACE;
	if (riid == IID_IUnknown || riid == IID_IDispatch || riid == id) {
		*ppv = static_cast<OposScanner_CCO::_IOPOSScannerEvents *>(this);
		AddRef();
		hr = S_OK;
	}
	return hr;
}

/* For reference counting. When IConnectionPoint.Advise() is called on
 * this, AddRef() is called. 
 */
IFACEMETHODIMP_(ULONG) 
ScannerSink::AddRef() {
	return ++ref;
}

/* For reference counting. When IConnectionPoint.UnAdvise() is called on
 * this, Release() is called and this is destroyed.
 */
IFACEMETHODIMP_(ULONG) 
ScannerSink::Release() {
	if (--ref == 0)
		delete this;
	return ref;
}

// IDispatch methods
IFACEMETHODIMP 
ScannerSink::GetTypeInfoCount(UINT *pctinfo) {
	*pctinfo = 0;
	return E_NOTIMPL;
}
IFACEMETHODIMP 
ScannerSink::GetTypeInfo(UINT itinfo,
	LCID lcid, ITypeInfo **iti)
{
	*iti = nullptr;
	return E_NOTIMPL;
}

/* Called by the COM object to retrieve information about the specific
 * properties or methods that this implements. Sets dispids with the 
 * relevant dispid of the name passed in through names.
 */
IFACEMETHODIMP 
ScannerSink::GetIDsOfNames(REFIID riid, LPOLESTR *names,
	UINT size, LCID lcid, DISPID *dispids)
{
	if (!wcscmp(names[0], to_wstr(ScannerEvent::StatusUpdate).c_str())) {
		dispids[0] = ScannerEvent::StatusUpdate;
	}
	else if (!wcscmp(names[0], to_wstr(ScannerEvent::DirectIO).c_str())) {
		dispids[0] = ScannerEvent::DirectIO;
	}
	else if (!wcscmp(names[0], to_wstr(ScannerEvent::Error).c_str())) {
		dispids[0] = ScannerEvent::Error;
	}
	else if (!wcscmp(names[0], to_wstr(ScannerEvent::Data).c_str())) {
		dispids[0] = ScannerEvent::Data;
	}
	else {
		dispids[0] = -1;
		return E_NOTIMPL;
	}

	return S_OK;
}

/* Called by the COM object when a scanner event occurs. The scanner event 
 * corresponds to a dispid determined when this.GetIDsOfNames() was called
 *
 */
IFACEMETHODIMP 
ScannerSink::Invoke(DISPID dispid, REFIID riid, LCID lcid,
	WORD flags, DISPPARAMS *dispparams, VARIANT *result,
	EXCEPINFO *exceptioninfo, UINT *argerr) 
{
	switch (dispid) {
	case ScannerEvent::Data:
		return DataEvent(dispparams->rgvarg[0].lVal);
	case ScannerEvent::DirectIO:
		return S_OK;
	case ScannerEvent::Error:
		return S_OK;
	case ScannerEvent::StatusUpdate:
		return S_OK;
	default:
		return S_OK;
	}
}

// _IOPOSScannerEvents methods
HRESULT
ScannerSink::DataEvent(long Status) {
	std::cout << "Data: " << scanner.ScanDataLabel << std::endl;
	// scanner.DataEventEnabled is set to false when a DataEvent is invoked
	// the program must set it back to true to continue recieving DataEvents
	scanner.DataEventEnabled = true;
	return S_OK;
}
HRESULT 
ScannerSink::DirectIOEvent(
	long EventNumber,
	long *pData,
	BSTR *pString)
{
	return S_OK;
}
HRESULT 
ScannerSink::ErrorEvent(
	long ResultCode,
	long ResultCodeExtended,
	long ErrorLocus,
	long *pErrorResponse)
{
	return S_OK;
}
HRESULT 
ScannerSink::StatusUpdateEvent(
	long Data)
{	
	return S_OK;
}

std::string 
ScannerSink::to_str(const ScannerEvent &e) {
	switch (e) {
	case ScannerEvent::Data:
		return "DataEvent";
	case ScannerEvent::DirectIO:
		return "DirectIOEvent";
	case ScannerEvent::Error:
		return "ErrorEvent";
	case ScannerEvent::StatusUpdate:
		return "StatusUpdateEvent";
	default:
		return "Unused";
	}
}

std::wstring
ScannerSink::to_wstr(const ScannerEvent &e) {
	switch (e) {
	case ScannerEvent::Data:
		return L"DataEvent";
	case ScannerEvent::DirectIO:
		return L"DirectIOEvent";
	case ScannerEvent::Error:
		return L"ErrorEvent";
	case ScannerEvent::StatusUpdate:
		return L"StatusUpdateEvent";
	default:
		return L"Unused";
	}
}