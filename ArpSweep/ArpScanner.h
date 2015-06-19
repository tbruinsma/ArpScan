/* Copyright (c) Constantine V. Sharlaimov   */
/* ArpScanner.h : Declaration of the CArpScanner */

#pragma once
//#include <windows.h>
#include "resource.h"       // main symbols

/****************************************/
#define ROL32(x,n) _lrotl((x), (n))
#define ROR32(x,n) _lrotr((x), (n))
#define BSWAP_32(x) (ROR32((x), 24) & 0x00FF00FF | ROR32((x), 8) & 0xFF00FF00)
#define BSWAP_16(x) (((WORD)(x) << 8) & 0xFF00 | ((WORD)(x) >> 8) & 0x00FF)

#define  E_INVALID_STARTIP                (1)
#define  E_INVALID_ENDIP                  (2)
#define  E_INVALID_ADAPTER                (3)
#define  E_SCAN_IN_PROGRESS               (4)
#define  E_WINPCAP_NOT_FOUND              (5)
#define  E_COULD_NOT_OPEN_ADAPTER         (6)
#define  E_INVALID_IP                     (7)
#define  E_INVALID_MAC                    (8)
#define	 E_INVALID_KEY					  (9)

/* The scan of N ip addresses will take N * MAX_ARP_SCAN_TIMEOUT / MAX_CONCURRENT_SCANS ms.
 * Increasing MAX_CONCURRENT_SCANS will speed up scanning but raise false OnScanHost count,
 * occuring when TCP/IP performs its own ARP request to a host being scaned by ARPScanner.
 */

#define  MAX_CONCURRENT_SCANS             (8)         /* 8 should be good */
#define  MAX_ARP_SCAN_TIMEOUT             (1000)      /* 1 sec = 1000 msec */

struct TARPScannerContext
{
   BOOL        requested;
   BOOL        replied;
   IN_ADDR     addr;
   DWORD       timeout;
};

#include "winpcap-wrapper.h"
/****************************************/

// IArpScanner
[
	object,
	uuid("12F0C3AF-5936-4601-A810-F0326F805952"),
	dual,	helpstring("IArpScanner Interface"),
	pointer_default(unique)
]
__interface IArpScanner : IDispatch
{
   [propget, id(1), helpstring("property StartIP")] HRESULT StartIP([out, retval] BSTR* pVal);
   [propput, id(1), helpstring("property StartIP")] HRESULT StartIP([in] BSTR newVal);
   [propget, id(2), helpstring("property EndIP")] HRESULT EndIP([out, retval] BSTR* pVal);
   [propput, id(2), helpstring("property EndIP")] HRESULT EndIP([in] BSTR newVal);
   [propput, id(12), helpstring("property Key")] HRESULT Key([in] LONG newVal);
   [propget, id(3), helpstring("property Adapter")] HRESULT Adapter([out, retval] LONG* pVal);
   [propput, id(3), helpstring("property Adapter")] HRESULT Adapter([in] LONG newVal);
   [propget, id(4), helpstring("property HostCount")] HRESULT HostCount([out, retval] LONG* pVal);
   [id(5), helpstring("method StartScan")] HRESULT StartScan(void);
   [propget, id(6), helpstring("property AdapterCount")] HRESULT AdapterCount([out, retval] LONG* pVal);
   [id(7), helpstring("method GetAdapterInfo")] HRESULT GetAdapterInfo([in] LONG adapter, [out] BSTR * name, [out] BSTR* description, [out] BSTR* ip, [out] BSTR* mac, [out, retval] BOOL * ok);
   [id(8), helpstring("method Scan")] HRESULT Scan([in] LONG adapter, [in] BSTR ip, [in] BSTR mac, [out,retval] LONG * response);
   [id(9), helpstring("method StopScan")] HRESULT StopScan(void);
   [id(10), helpstring("method LookupAdapterName")] HRESULT LookupAdapterName([in] BSTR AdapterName, [out,retval] LONG* Adapter);
   [propget, id(11), helpstring("property ScannedHostCount")] HRESULT ScannedHostCount([out, retval] LONG* pVal);
};


// _IArpScannerEvents
[
	dispinterface,
	uuid("59753AB6-4CD0-40D2-9EFC-FA452CBF2051"),
	helpstring("_IArpScannerEvents Interface")
]
__interface _IArpScannerEvents
{
   [id(1), helpstring("method OnScanHost")] HRESULT OnScanHost([in] BSTR ip, BSTR mac, LONG response_time);
   [id(2), helpstring("method OnScanStart")] HRESULT OnScanStart(void);
   [id(3), helpstring("method OnStartStop")] HRESULT OnStartStop(void);
   [id(4), helpstring("method OnScanEvent")] HRESULT OnScanError([in] LONG code);
};


// CArpScanner

[
	coclass,
	threading("free"),
	support_error_info("IArpScanner"),
	event_source("com"),
	vi_progid("ArpSweep.ArpScanner"),
	progid("ArpSweep.ArpScanner.1"),
	version(1.0),
	uuid("F6CDAC27-27C7-4030-BAC7-E3E7143DDE0B"),
	helpstring("ArpScanner Class")
]
class ATL_NO_VTABLE CArpScanner : public IArpScanner
{
public:
	CArpScanner()
	{
      /* variable initialization */
      this->sci_Adapter = -1;
      this->sci_StartIP.S_un.S_addr = INADDR_NONE;
      this->sci_EndIP.S_un.S_addr = INADDR_NONE;
      this->scan_thread = NULL;

      /* initialize concurrent scan contexts */
      for (int i = 0; i < MAX_CONCURRENT_SCANS; i++) this->scan_ctx[i].timeout = 0; /* passed long-long ago */

      /* load PCAP */
      this->pcap.LoadPCAP();
	}

   ~CArpScanner()
   {
      /* Terminate the thread (won't cause any side effects) */
      TerminateThread(this->scan_thread, 0);
      CloseHandle(this->scan_thread);
   }

	__event __interface _IArpScannerEvents;

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
	}

public:

   STDMETHOD(get_StartIP)(BSTR* pVal);
   STDMETHOD(put_StartIP)(BSTR newVal);
   STDMETHOD(get_EndIP)(BSTR* pVal);
   STDMETHOD(put_EndIP)(BSTR newVal);
   STDMETHOD(put_Key)(LONG newVal);
   STDMETHOD(get_Adapter)(LONG* pVal);
   STDMETHOD(put_Adapter)(LONG newVal);
   STDMETHOD(get_HostCount)(LONG* pVal);
   STDMETHOD(get_AdapterCount)(LONG* pVal);
   STDMETHOD(get_ScannedHostCount)(LONG* pVal);
   STDMETHOD(GetAdapterInfo)(LONG adapter, BSTR * name, BSTR* description, BSTR* ip, BSTR* mac, BOOL * ok);
   STDMETHOD(StopScan)(void);
   STDMETHOD(StartScan)(void);
   STDMETHOD(Scan)(LONG adapter, BSTR ip, BSTR mac, LONG * response);
   STDMETHOD(LookupAdapterName)(BSTR AdapterName, LONG* Adapter);

private:
   IN_ADDR              sci_StartIP;
   IN_ADDR              sci_EndIP;
   LONG                 sci_Adapter;
   LONG					sci_Key;
   HANDLE               scan_thread;
   TARPScannerContext   scan_ctx[MAX_CONCURRENT_SCANS];
   TWinPCAP             pcap;
   LONG volatile        scanned_host_count;
   bool volatile        terminate_scan;

   bool ScansActive(VOID) { for (int i = 0; i < MAX_CONCURRENT_SCANS; i++) if (scan_ctx[i].requested) return TRUE; return FALSE; }
   bool CheckAdapterNumber(long adapter);

   /* friend thread function - to access private fields */
   friend DWORD CALLBACK ArpScanThread(CArpScanner * cls);
   friend void CheckPCAPPackets(CArpScanner * cls, TWinPCAP * pcap, IN_ADDR * adapter_ip);
   friend LONG ArpScanSingle(CArpScanner * cls, int adapter, IN_ADDR ip, unsigned char mac[6]);
};

/**********************************************************************************************************/
//DWORD CALLBACK ArpScanThread(CArpScanner * cls);
/**********************************************************************************************************/
