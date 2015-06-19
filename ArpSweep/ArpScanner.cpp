/* Copyright (c) Constantine V. Sharlaimov   */
/* ArpScanner.cpp : Implementation of CArpScanner */

#include "stdafx.h"
#include "ArpScanner.h"
#include ".\arpscanner.h"


/* CArpScanner */
STDMETHODIMP CArpScanner::get_StartIP(BSTR* pVal)
{
   CComBSTR tmp(inet_ntoa(this->sci_StartIP));
   *pVal = tmp.Copy();
   return S_OK;
}

STDMETHODIMP CArpScanner::put_StartIP(BSTR newVal)
{
   /* can modify parameters only when scan is inactive */
   if (this->scan_thread == NULL)
   {
      CComBSTR tmp(newVal);
      CW2CT szIP(tmp);
      ULONG raw_ip = inet_addr(szIP);

      if (raw_ip != INADDR_NONE)
      {
         this->sci_StartIP.S_un.S_addr = raw_ip;
         return S_OK;
      }
      else
      {
         this->sci_StartIP.S_un.S_addr = INADDR_NONE;
         __raise OnScanError(E_INVALID_STARTIP);
         return S_FALSE;
      }
   }
   else
   {
      __raise OnScanError(E_SCAN_IN_PROGRESS);
      return S_FALSE;
   }
}

STDMETHODIMP CArpScanner::get_EndIP(BSTR* pVal)
{
   CComBSTR tmp(inet_ntoa(this->sci_EndIP));
   *pVal = tmp.Copy();
   return S_OK;
}

STDMETHODIMP CArpScanner::put_EndIP(BSTR newVal)
{
   /* can modify parameters only when scan is inactive */
   if (this->scan_thread == NULL)
   {
      CComBSTR tmp(newVal);
      CW2CT szIP(tmp);
      ULONG raw_ip = inet_addr(szIP);

      if (raw_ip != INADDR_NONE)
      {
         this->sci_EndIP.S_un.S_addr = raw_ip;
         return S_OK;
      }
      else
      {
         this->sci_EndIP.S_un.S_addr = INADDR_NONE;
         __raise OnScanError(E_INVALID_ENDIP);
         return S_FALSE;
      }
   }
   else
   {
      __raise OnScanError(E_SCAN_IN_PROGRESS);
      return S_FALSE;
   }
}

STDMETHODIMP CArpScanner::put_Key(LONG newVal)
{
   this->sci_Key = newVal;
   return S_OK;
}

STDMETHODIMP CArpScanner::get_Adapter(LONG* pVal)
{
   *pVal = this->sci_Adapter;
   return S_OK;
}

STDMETHODIMP CArpScanner::put_Adapter(LONG newVal)
{
   /* can modify parameters only when scan is inactive */
   if (this->scan_thread == NULL)
   {
      if (this->CheckAdapterNumber(newVal))
      {
         this->sci_Adapter = newVal;
         return S_OK;
      }
      else
      {
         this->sci_Adapter = -1;
         __raise OnScanError(E_INVALID_ADAPTER); 
         return S_FALSE;
      }
   }
   else
   {
      __raise OnScanError(E_SCAN_IN_PROGRESS);
      return S_FALSE;
   }
}

STDMETHODIMP CArpScanner::get_HostCount(LONG* pVal)
{
   if (this->sci_StartIP.S_un.S_addr == INADDR_NONE) { __raise OnScanError(E_INVALID_STARTIP); *pVal = 0; return S_FALSE; }
   if (this->sci_EndIP.S_un.S_addr == INADDR_NONE) { __raise OnScanError(E_INVALID_ENDIP); *pVal = 0; return S_FALSE; }

   *pVal = BSWAP_32(this->sci_EndIP.S_un.S_addr) - BSWAP_32(this->sci_StartIP.S_un.S_addr) + 1;
   if (*pVal < 0) *pVal = 0;

   return S_OK;
}

STDMETHODIMP CArpScanner::StartScan(void)
{
   DWORD scratch;

   /* check if scan is already started, check for bad parameters */
   if (this->sci_Key!= 328760823) { __raise OnScanError(E_INVALID_KEY); return S_FALSE; }
   if (this->scan_thread != NULL) { __raise OnScanError(E_SCAN_IN_PROGRESS); return S_FALSE; }
   if (this->sci_StartIP.S_un.S_addr == INADDR_NONE) { __raise OnScanError(E_INVALID_STARTIP); return S_FALSE; }
   if (this->sci_EndIP.S_un.S_addr == INADDR_NONE) { __raise OnScanError(E_INVALID_ENDIP); return S_FALSE; }
   if (!this->CheckAdapterNumber(this->sci_Adapter)) { __raise OnScanError(E_INVALID_ADAPTER); return S_FALSE; }

   /* Start thread */
   this->terminate_scan = false;
   if ((this->scan_thread = CreateThread(NULL, 0, (LPTHREAD_START_ROUTINE)ArpScanThread, (LPVOID)this, 0, &scratch)) != NULL)
      return S_OK;
   else
      return S_FALSE;
}

STDMETHODIMP CArpScanner::get_AdapterCount(LONG* pVal)
{
   if (this->pcap.LoadPCAP())
   {
      *pVal = this->pcap.GetAdapterCount();
      return S_OK;
   }
   else
   {
      __raise this->OnScanError(E_WINPCAP_NOT_FOUND);
      return S_FALSE;
   }   
}

STDMETHODIMP CArpScanner::GetAdapterInfo(LONG adapter, BSTR * name, BSTR* description, BSTR* ip, BSTR* mac, BOOL * ok)
{
   if (this->pcap.LoadPCAP())
   {
      if (CheckAdapterNumber(adapter))
      {
         IN_ADDR  raw_ip;
         unsigned char raw_mac[6];
         char mac_txt[18];

         this->pcap.GetAdapterIP(adapter, &raw_ip, NULL);
         this->pcap.GetAdapterMAC(adapter, raw_mac);
         MAC2TXT(mac_txt, raw_mac);

         CComBSTR adapter_ip(inet_ntoa(raw_ip));
         CComBSTR adapter_mac(mac_txt);
         CComBSTR adapter_descr(this->pcap.GetAdapterDescription(adapter));
         CComBSTR adapter_name(this->pcap.GetAdapterName(adapter));

         *name = adapter_name.Copy();
         *description = adapter_descr.Copy();
         *ip = adapter_ip.Copy();
         *mac = adapter_mac.Copy();

         *ok = true;
         return S_OK;
      }
      else
      {
         *ok = false;
         __raise this->OnScanError(E_INVALID_ADAPTER);
         return S_FALSE;
      }
   }
   else
   {
      *ok = false;
      __raise this->OnScanError(E_WINPCAP_NOT_FOUND);
      return S_FALSE;
   }   
}

STDMETHODIMP CArpScanner::Scan(LONG adapter, BSTR ip, BSTR mac, LONG * response)
{
   unsigned char  dst_mac[6];
   IN_ADDR        dst_ip;

   *response = -1;

   /* check if scan is already started, check for bad parameters */
   if (this->scan_thread != NULL) { __raise OnScanError(E_SCAN_IN_PROGRESS); return S_FALSE; }
   if (!this->CheckAdapterNumber(adapter)) { __raise OnScanError(E_INVALID_ADAPTER); return S_FALSE; }

   CComBSTR _ip(ip);
   CComBSTR _mac(mac);

   CW2CT szIP(_ip);
   CW2CT szMAC(_mac);

   dst_ip.S_un.S_addr = inet_addr(szIP);
   if (dst_ip.S_un.S_addr == INADDR_NONE) { __raise OnScanError(E_INVALID_IP); return S_FALSE; }
   if (!TXT2MAC(dst_mac, szMAC)) { __raise OnScanError(E_INVALID_MAC); return S_FALSE; }

   /* perform actual scan (blocking) */
   *response = ArpScanSingle(this, adapter, dst_ip, dst_mac);

   return S_OK;
}

STDMETHODIMP CArpScanner::StopScan(void)
{
   /* just a signal, other will be handled by scanning thread */
   this->terminate_scan = true;
   return S_OK;
}

STDMETHODIMP CArpScanner::LookupAdapterName(BSTR AdapterName, LONG* Adapter)
{
   if (this->pcap.LoadPCAP())
   {
      CComBSTR name(AdapterName);
      INT i, n;

      *Adapter = -1;

      for (i = 0, n = this->pcap.GetAdapterCount(); i < n; i++)
      {
         CComBSTR lookup(this->pcap.GetAdapterName(i));
         if (lookup == name) 
         {
            *Adapter = i;
            break;
         }
      }

      return S_OK;
   }
   else
   {
      *Adapter = -1;
      __raise this->OnScanError(E_WINPCAP_NOT_FOUND);
      return S_FALSE;
   }   
}

STDMETHODIMP CArpScanner::get_ScannedHostCount(LONG* pVal)
{
   if (this->scan_thread != NULL)
   {
      *pVal = this->scanned_host_count;
      return S_OK;
   }
   else
   {
      *pVal = 0;
      return S_OK;
   }
}

bool CArpScanner::CheckAdapterNumber(LONG adapter)
{
   if (this->pcap.LoadPCAP())
   {
      if (adapter >= 0 && adapter < this->pcap.GetAdapterCount())
      {
         return true;
      }
      else
      {
         return false;
      }
   }
   else
   {
      __raise this->OnScanError(E_WINPCAP_NOT_FOUND);
      return false;;
   }   
}
