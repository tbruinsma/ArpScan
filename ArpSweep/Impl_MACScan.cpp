/* Copyright (c) Constantine V. Sharlaimov   */
/* Implementation of the scanner (thread)    */

#include "stdafx.h"
#include "ArpScanner.h"
#include ".\arpscanner.h"

/*******************************************************************************************/
/** Begin of WinPCAP Wrapper class *********************************************************/
/*******************************************************************************************/
/* jshadow: constructor - zero all */
TWinPCAP::TWinPCAP(void)
{
   this->h_pcap_lib = NULL;
   this->alldevs = NULL;
   this->adapter = NULL;
}

/* jshadow: cleanup in case programmer forgot something */
TWinPCAP::~TWinPCAP(void)
{
   if (this->IsLoaded())
   {
      if (this->adapter) this->pcap_close(this->adapter);
      if (this->alldevs) this->pcap_freealldevs(this->alldevs);
      FreeLibrary(this->h_pcap_lib);
   }
}

/* jshadow: load library, import required functions, get adapter list */
bool TWinPCAP::LoadPCAP(void)
{
   if (!this->IsLoaded())
   {
      if ((this->h_pcap_lib = LoadLibrary("wpcap.dll")) != NULL)
      {
         this->pcap_open =        (pcap_t * (__cdecl *)(const char *, int, int, int, struct pcap_rmtauth *, char *)) GetProcAddress(this->h_pcap_lib, "pcap_open");
         this->pcap_close =       (void (__cdecl *)(pcap_t *)) GetProcAddress(this->h_pcap_lib, "pcap_close");
         this->pcap_datalink =    (int (__cdecl *)(pcap_t *)) GetProcAddress(this->h_pcap_lib, "pcap_datalink");
         this->pcap_next_ex =     (int (__cdecl *)(pcap_t *, struct pcap_pkthdr **, const u_char **)) GetProcAddress(this->h_pcap_lib, "pcap_next_ex");
         this->pcap_findalldevs = (int (__cdecl *)(pcap_if_t **, char *)) GetProcAddress(this->h_pcap_lib, "pcap_findalldevs");
         this->pcap_freealldevs = (void (__cdecl *)(pcap_if_t *)) GetProcAddress(this->h_pcap_lib, "pcap_freealldevs");
         this->pcap_sendpacket =  (int (__cdecl *)(pcap_t *, const u_char *, int)) GetProcAddress(this->h_pcap_lib, "pcap_sendpacket");

         if (this->pcap_open && this->pcap_close && this->pcap_datalink && this->pcap_next_ex &&
            this->pcap_findalldevs && this->pcap_freealldevs && this->pcap_sendpacket)
         {
            if (this->GetAllDevs())
            {
               return true;
            }
         }

         FreeLibrary(this->h_pcap_lib);
         this->h_pcap_lib = NULL;
      }

      return false;
   }
   else
   {
      return true;
   }
}

/* count all winpcap adapters */
int TWinPCAP::GetAdapterCount(void)
{
   int count;
   pcap_if_t * d;

   for (count = 0, d = this->alldevs; d; d = d->next, count++);

   return count;
}

/* helper - returns a pointer to an adapter */
pcap_if_t * TWinPCAP::GetAdapter(int adapter)
{
   int count;
   pcap_if_t * d;

   for (count = 0, d = this->alldevs; d; d = d->next, count++)
   {
      if (count == adapter) return d;
   }

   return NULL;
}

/* returns adapter name */
char * TWinPCAP::GetAdapterName(int adapter)
{
   pcap_if_t * d = GetAdapter(adapter);
   return (d) ? d->name : NULL;
}

/* returns adapter description */
char * TWinPCAP::GetAdapterDescription(int adapter)
{
   pcap_if_t * d = GetAdapter(adapter);
   return (d) ? d->description : NULL;
}

/* jShadow: black magic!  here we merge IP Helper with WinPCAP - we go through all 
 * IP addresses via IP Helper and check if some of them matches one of the addresses
 * of WinPCAP adapter address list. if a match is found we get MAC via IP Helper
 */
bool TWinPCAP::GetAdapterMAC(int adapter, unsigned char * mac)
{
   bool ok = false;

   /* Preparation: PCAP */
   pcap_if_t * d = GetAdapter(adapter);
   if (!d) return false;

   /* Preparation: IPHlpAPI */
   PMIB_IPADDRTABLE mib_ip_table = NULL;
   ULONG mib_ip_table_size = 0;
   
   /* walk through adapters via IPHelper API*/
   GetIpAddrTable(mib_ip_table, &mib_ip_table_size, FALSE);
   if (mib_ip_table_size != 0)
   {
      mib_ip_table = (PMIB_IPADDRTABLE) new unsigned char[mib_ip_table_size];
      if (mib_ip_table)
      {
         if (GetIpAddrTable(mib_ip_table, &mib_ip_table_size, FALSE) == NO_ERROR)
         {
            for (unsigned int i = 0; i < mib_ip_table->dwNumEntries && !ok; i++)
            {

               /* walk through adapter addresses */
               for (pcap_addr_t * addr = d->addresses; addr && !ok; addr = addr->next)
               {
                  /* we are interested only in AF_INET (IPv4) 'cause ARP is defined in IPv4 stack */
                  if (addr->addr->sa_family == AF_INET && ((struct sockaddr_in *)addr->addr)->sin_addr.S_un.S_addr == mib_ip_table->table[i].dwAddr)
                  {
                     /* gotcha - fetch HW address! */

                     MIB_IFROW mib_if_entry;

                     mib_if_entry.dwIndex = mib_ip_table->table[i].dwIndex;
                     if (GetIfEntry(&mib_if_entry) == NO_ERROR)
                     {
                        ok = true;
                        memcpy(mac, mib_if_entry.bPhysAddr, 6);
                     } /* if */
                  } /* if */
               } /* for on winpcap adapter list */

            } /* for each of IP address matches */
         } /* if GetIpAddrTable */

         /* cleanup */
         delete [] mib_ip_table;
      }
   }

   /* hack: we fill mac address with FF-s if we failed to find it. this
    * workaround is needed to avoid 3 more lines of code in send_packet -
    * if we failed to figure out mac address, we'll send from broadcast 
    * address (we'll most likely get a responce in that case). To speed 
    * up send_packet code we move invariant MAC workaroud here.
    */

   if (!ok)
   {
      memset(mac, 0xff, 6);
   }

   return ok;
}

/* returns true if opened adapter is ethernet-compatible (WiFi counts) */
bool TWinPCAP::IsEthernet(void)
{
   /* error: adapter not open */
   if (this->adapter == NULL) return false;
   
   return (this->pcap_datalink(this->adapter) == DLT_EN10MB);
}

/* opens specified adapter, caches its mac address */
bool TWinPCAP::OpenAdapter(int adapter)
{
   if (this->adapter != NULL) return false;  /* already open */

   char * adapter_name = this->GetAdapterName(adapter);

   if (adapter_name)
   {
      /* adapter found = open it */
      this->adapter = this->pcap_open(adapter_name, PCAP_MAX_PACKET_SIZE, PCAP_OPENFLAG_PROMISCUOUS, 20, NULL, this->errbuf);

      if (this->adapter)
      {
         /* ok, adapter opened - get its mac address and cache it for future use 
          * we don't care if GetAdapterMAC will fail due to a tiny workaround in 
          * GetAdapterMAC (see code for more description) */
         this->GetAdapterMAC(adapter, this->adapter_mac);

         return true;
      }
   }

   return false;
}

/* closes adapter */
bool TWinPCAP::CloseAdapter(void)
{
   if (this->adapter)
   {
      this->pcap_close(this->adapter);
      this->adapter = NULL;
      return true;
   }
   else
   {
      return false;
   }
}

/* pcap_sendpacket wrapper, some easiness of use added */
bool TWinPCAP::SendPacket(unsigned char * src_mac, unsigned char * dst_mac, int proto, unsigned char * payload, unsigned int payload_size)
{
   unsigned char pkt_data[PCAP_MAX_PACKET_SIZE];
   TEthernetPacket * pkt_header = (TEthernetPacket *)pkt_data;

   /* if src-mac not specified - use cached mac of opened adapter */
   if (src_mac) 
      memcpy(pkt_header->h_source, src_mac, 6);
   else
      memcpy(pkt_header->h_source, this->adapter_mac, 6);

   /* if dst-mac is not specified - broadcast the packet */
   if (dst_mac) 
      memcpy(pkt_header->h_dest, dst_mac, 6);
   else
      memset(pkt_header->h_dest, 0xff, 6);

   /* protocol */
   pkt_header->h_proto = BSWAP_16(proto);

   /* payload (if no payload specified - fill with zeroes) */
   if (payload)
      memcpy(pkt_data + sizeof(TEthernetPacket), payload, payload_size);
   else
      memset(pkt_data + sizeof(TEthernetPacket), 0x00, payload_size);

   /* do an actual send */
   if (this->pcap_sendpacket(this->adapter, pkt_data, sizeof(TEthernetPacket) + payload_size) != 0)
      return false;
   else
      return true;
}

/* receives an ethernet packet, parses it */
unsigned int TWinPCAP::ReadPacket(unsigned char * src_mac, unsigned char * dst_mac, int * proto, unsigned char * payload , unsigned int buffer_size)
{
   struct pcap_pkthdr * header;
   TEthernetPacket * pkt_data;

   if (this->pcap_next_ex(this->adapter, &header, (const u_char **)&pkt_data) == 1)
   {
      if (src_mac) memcpy(src_mac, pkt_data->h_source, 6);
      if (dst_mac) memcpy(dst_mac, pkt_data->h_dest, 6);
      if (proto) *proto = BSWAP_16(pkt_data->h_proto);
      if (payload) memcpy(payload, ((u_char *)pkt_data) + sizeof(TEthernetPacket), min(header->caplen - sizeof(TEthernetPacket), buffer_size));

      return header->caplen;
   }
   else
   {
      return 0;
   }
}

/* jShadow: an internal of the class - wrapper around pcap_findalldevs */
bool TWinPCAP::GetAllDevs(void)
{
   if (this->alldevs == NULL)
      return (this->pcap_findalldevs(&this->alldevs, this->errbuf) != -1);
   else
      return true;
}

/* jshadow: dumb - we get the first IP address (if it exists) */
bool TWinPCAP::GetAdapterIP(int adapter, IN_ADDR * ip, IN_ADDR * mask)
{
   /* Preparation: PCAP */
   pcap_if_t * d = GetAdapter(adapter);
   if (!d) return false;

   /* look for IP address */
   for (pcap_addr_t * addr = d->addresses; addr; addr = addr->next)
   {
      /* we are interested only in AF_INET (IPv4) 'cause ARP is defined in IPv4 stack */
      if (addr->addr->sa_family == AF_INET)
      {
         if (ip) ip->S_un.S_addr = ((struct sockaddr_in *)addr->addr)->sin_addr.S_un.S_addr;
         if (mask) mask->S_un.S_addr = ((struct sockaddr_in *)addr->netmask)->sin_addr.S_un.S_addr;
         return true;
      }
   }

   /* not found */
   return false;
}
/*******************************************************************************************/
/** End of WinPCAP Wrapper class ***********************************************************/
/*******************************************************************************************/

/*******************************************************************************************/
/** Begin of Ethernet MAC address helpers **************************************************/
/*******************************************************************************************/
#pragma pack(1)
struct TIPv4ARP
{
   unsigned __int16  hw_type;    /* 0001 for ethernet */
   unsigned __int16  proto;      /* 0800 for IPv4 */
   unsigned __int8   hlen;       /* 06 for ethernet */
   unsigned __int8   plen;       /* 04 for IPv4 */
   unsigned __int16  opcode;     /* 0001 for ARP request and 0002 for ARP response */
   unsigned __int8   src_mac[6];
   unsigned __int32  src_ip;
   unsigned __int8   dst_mac[6];
   unsigned __int32  dst_ip;
};
#pragma pack()

void ARP_PrepareRequest(TIPv4ARP * pkt, unsigned char * src_mac, IN_ADDR src_ip, IN_ADDR dst_ip)
{
   pkt->hw_type = 0x0100;  /* in network byte order */
   pkt->proto = 0x0008;
   pkt->hlen = 6;
   pkt->plen = 4;
   pkt->opcode = 0x0100;
   pkt->src_ip = src_ip.S_un.S_addr;
   pkt->dst_ip = dst_ip.S_un.S_addr;
   
   memset(pkt->dst_mac, 0x00, 6);
   memcpy(pkt->src_mac, src_mac, 6);
}

bool ARP_ParseResponse(TIPv4ARP * pkt, unsigned char * src_mac, IN_ADDR * src_ip, unsigned char * dst_mac, IN_ADDR * dst_ip)
{
   if (pkt->hw_type != 0x0100) return false;
   if (pkt->proto != 0x0008) return false;
   if (pkt->hlen != 6) return false;
   if (pkt->plen != 4) return false;
   if (pkt->opcode != 0x0200) return false;

   /* get info */
   if (src_mac) memcpy(src_mac, pkt->src_mac, 6);
   if (src_ip) src_ip->S_un.S_addr = pkt->src_ip;

   if (dst_mac) memcpy(dst_mac, pkt->dst_mac, 6);
   if (dst_ip) dst_ip->S_un.S_addr = pkt->dst_ip;

   return true;
}

/* text is 6*2 + 5 + 1 = 18 bytes long, mac is 6 bytes long */
void MAC2TXT(char * text, unsigned char * mac)
{
   sprintf(text, "%02x-%02x-%02x-%02x-%02x-%02x", mac[0], mac[1], mac[2], mac[3], mac[4], mac[5]);
}

int hex2byte(char * s)
{
   unsigned char res = 0;
   int i;

   for (i = 0; i < 2; i++)
   {
      res <<= 4;

      if (*s >= '0' && *s <= '9')
         res += (*s - '0');
      else if (*s >= 'A' && *s <= 'F')
         res += (*s - 'A') + 10;
      else if (*s >= 'a' && *s <= 'f')
         res += (*s - 'a') + 10;
      else return -1;

      s++;
   }

   return res;
}

/* text is 6*2 + 5 + 1 = 18 bytes long, mac is 6 bytes long */
bool TXT2MAC(unsigned char * mac, char * text)
{
   char  tmp[18];
   int   maci[6];

   /* create a temporary storage */
   strncpy(tmp, text, 17);
   tmp[17] = '\0';

   /*           1111111 */
   /* 01234567890123456 */
   /* 00-00-00-00-00-00 */

   for (int i = 0; i < 5; i++) 
   {
      if (tmp[2 + i * 3] != '-' && tmp[2 + i * 3] != ':') return false;
      tmp[2 + i * 3] = '\0';
   }

   for (int i = 0; i < 6; i++) maci[i] = hex2byte(&tmp[0 + i * 3]); 

   if (maci[0] >= 0 && maci[1] >= 0 && maci[2] >= 0 && maci[3] >= 0 && maci[4] >= 0 && maci[5] >= 0)
   {
      for (int i = 0; i < 6; i++) mac[i] = maci[i];
      return true;
   }
   else
   {
      return false;
   }
}

/*******************************************************************************************/
/** End of Ethernet MAC address helpers ****************************************************/
/*******************************************************************************************/

/*******************************************************************************************/
/** Begin of broadcast ARP Scanner *********************************************************/
/*******************************************************************************************/
void CheckPCAPPackets(CArpScanner * cls, TWinPCAP * pcap, IN_ADDR * adapter_ip)
{
   int proto;
   unsigned int size;
   unsigned char packet[PCAP_MAX_PACKET_SIZE];

   /* handle received packets */
   if ((size = pcap->ReadPacket(NULL, NULL, &proto, packet, PCAP_MAX_PACKET_SIZE)) >= sizeof(TIPv4ARP))
   {
      if (proto == 0x0806) /* ARP */
      {
         unsigned char arp_src_mac[6];
         IN_ADDR arp_src_ip;
         IN_ADDR arp_dst_ip;

         /* if ARP response is valid and parsed */
         if (ARP_ParseResponse((TIPv4ARP *)packet, arp_src_mac, &arp_src_ip, NULL, &arp_dst_ip))
         {
            /* verify that the response is sent to our IP address */
            if (arp_dst_ip.S_un.S_addr == adapter_ip->S_un.S_addr)
            {
               /* Verify context */
               for (int i = 0; (i < MAX_CONCURRENT_SCANS); i++)
               {
                  /* not a timed out context and send a request to IP */
                  if ((cls->scan_ctx[i].timeout >= GetTickCount()) && (cls->scan_ctx[i].addr.S_un.S_addr == arp_src_ip.S_un.S_addr))
                  {
                     /* mark context as replied, do not mark as not requested
                      * for we can receive MULTIPLE replies (ip spooffing)
                      */
                     cls->scan_ctx[i].replied = true;

                     /* found a context - raise an event */
                     char arp_src_mac_txt[18];
                     MAC2TXT(arp_src_mac_txt, arp_src_mac);

                     /* Raise an event */
                     CComBSTR tmp_ip(inet_ntoa(arp_src_ip));
                     CComBSTR tmp_mac(arp_src_mac_txt);
                     __raise cls->OnScanHost(tmp_ip.Copy(), tmp_mac.Copy(), GetTickCount() + MAX_ARP_SCAN_TIMEOUT - cls->scan_ctx[i].timeout);

                     return;
                  }
               } /* context verification loop */
            } /* destination IP = our IP */
         } /* ARP parse */
      } /* proto verification */
   } /* cls->pcap.ReadPacket */
}

DWORD CALLBACK ArpScanThread(CArpScanner * cls)
{
   /* Initialize WinPCAP */
   //TWinPCAP pcap;

   if (cls->pcap.LoadPCAP())
   {
      /* Get adapter parameters */
      IN_ADDR adapter_ip;
      unsigned char adapter_mac[6];

      /* Get IP and MAC */
      cls->pcap.GetAdapterIP(cls->sci_Adapter, &adapter_ip, NULL);
      cls->pcap.GetAdapterMAC(cls->sci_Adapter, adapter_mac);

      /* Clear all contexts */
      for (int i = 0; i < MAX_CONCURRENT_SCANS; i++)
      {
         cls->scan_ctx[i].addr.S_un.S_addr = 0; /* IP address we send packet for */
         cls->scan_ctx[i].replied = false;      /* was out request replied at least one time */
         cls->scan_ctx[i].timeout = 0;          /* when will this request time out */
         cls->scan_ctx[i].requested = false;    /* the request was sent and waiting for response */
      }

      /* Set starting ip address */
      IN_ADDR  current_ip = cls->sci_StartIP;

      /* start WinPCAP adapter capturing */
      if (cls->pcap.OpenAdapter(cls->sci_Adapter))
      {
         /* Scan started */
         cls->scanned_host_count = 0;
         __raise cls->OnScanStart();

         /* prepare */
         DWORD last_arp_sent = 0;

         /* loop while there are waiting scans and ending IP is not reached */
         do
         {
            /* check for received packets, parse them, verify context and raise OnScanHost events */
            CheckPCAPPackets(cls, &cls->pcap, &adapter_ip);

            /* check for context freeing and allocate next IP */
            for (int i = 0; i < MAX_CONCURRENT_SCANS; i++)
            {
               /* check for timeout */
               if (cls->scan_ctx[i].timeout < GetTickCount())
               {
                  /* requested but unreplied */
                  if (cls->scan_ctx[i].requested && !cls->scan_ctx[i].replied)
                  {
                     CComBSTR tmp_ip(inet_ntoa(cls->scan_ctx[i].addr));
                     cls->scanned_host_count++;

                     /* 1 = functional, 0 = demo */
#if 1
                     /* Actual - signal a timeout */
                     CComBSTR tmp_mac("");
                     __raise cls->OnScanHost(tmp_ip.Copy(), tmp_mac.Copy(), -1);
#else
                     /* Demo - random results */
                     if (rand() % 2)
                     {
                        unsigned char t_mac[6];
                        char          t_mac_txt[18];
                        for (int i = 0; i < 6; t_mac[i++] = rand() & 0xFF);
                        MAC2TXT(t_mac_txt, t_mac);
                        CComBSTR tmp_mac(t_mac_txt);
                        __raise cls->OnScanHost(tmp_ip.Copy(), tmp_mac.Copy(), rand() % 1000);
                     }
                     else
                     {
                        CComBSTR tmp_mac("");
                        __raise cls->OnScanHost(tmp_ip.Copy(), tmp_mac.Copy(), -1);
                     }
#endif
                  }
                  else if (cls->scan_ctx[i].requested && cls->scan_ctx[i].replied)
                  {
                     /* just count a host - events were raised in CheckPCAPPackets */
                     cls->scanned_host_count++;
                  }

                  /* in case of timeout mark the context as not requested */
                  cls->scan_ctx[i].requested = false;
               }

               /* is the context free so we can use it (and last request was sent enough time ago ? */
               if ((!cls->scan_ctx[i].requested) && 
                   (BSWAP_32(current_ip.S_un.S_addr) <= BSWAP_32(cls->sci_EndIP.S_un.S_addr)) && 
                   (GetTickCount() > (last_arp_sent + (MAX_ARP_SCAN_TIMEOUT / MAX_CONCURRENT_SCANS))))
               {
                  /* initialize a context with new address and timeout */
                  cls->scan_ctx[i].timeout = GetTickCount() + MAX_ARP_SCAN_TIMEOUT;
                  cls->scan_ctx[i].addr.S_un.S_addr = current_ip.S_un.S_addr;
                  cls->scan_ctx[i].requested = true;
                  cls->scan_ctx[i].replied = false;

                  /* craft an ARP Request packet and Send it to the network */
                  TIPv4ARP arp_pkt;
                  ARP_PrepareRequest(&arp_pkt, adapter_mac, adapter_ip, current_ip);
                  cls->pcap.SendPacket(NULL, NULL, 0x0806, (unsigned char *)&arp_pkt, sizeof(arp_pkt));
                  last_arp_sent = GetTickCount();

                  /* jump to next address */
                  current_ip.S_un.S_addr = BSWAP_32(BSWAP_32(current_ip.S_un.S_addr) + 1);
               }
            }

            /* Give up some CPU time */
            Sleep(1);
         } while (cls->ScansActive() && !cls->terminate_scan);

         /* Scan ended */
         __raise cls->OnStartStop();

         /* close adapter */
         cls->pcap.CloseAdapter();
      }
      else
      {
         /* error opening adapter */
         __raise cls->OnScanError(E_COULD_NOT_OPEN_ADAPTER);
      }
   }
   else
   {
      /* error loading WinPCAP */
      __raise cls->OnScanError(E_WINPCAP_NOT_FOUND);
   }

   /* Perform thread cleanup */
   CloseHandle(cls->scan_thread);
   cls->scan_thread = NULL;

   return 0;
}

/* jshadow: ping a single IP+MAC pair and return timeout */
LONG ArpScanSingle(CArpScanner * cls, int adapter, IN_ADDR ip, unsigned char mac[6])
{
   LONG response_time = -1;

   cls->scan_thread = INVALID_HANDLE_VALUE;        /* emulate running thread */

   if (cls->pcap.LoadPCAP())
   {
      /* Get adapter parameters */
      IN_ADDR adapter_ip;
      unsigned char adapter_mac[6];

      /* Get IP and MAC */
      cls->pcap.GetAdapterIP(adapter, &adapter_ip, NULL);
      cls->pcap.GetAdapterMAC(adapter, adapter_mac);

      /* start WinPCAP adapter capturing */
      if (cls->pcap.OpenAdapter(adapter))
      {
         /* craft and send ARP request */
         TIPv4ARP arp_pkt;

         /* prepare as a broadcast */
         ARP_PrepareRequest(&arp_pkt, adapter_mac, adapter_ip, ip);

         /* send directional (unicast to specified MAC) 
          * This is not RFC826 but it seem to work fine
          */
         cls->pcap.SendPacket(NULL, mac, 0x0806, (unsigned char *)&arp_pkt, sizeof(arp_pkt));

         /* wait for timeout */
         for (DWORD timeout = GetTickCount() + MAX_ARP_SCAN_TIMEOUT; GetTickCount() < timeout && response_time == -1; )
         {
            /* read packets and wait for response */

            int proto;
            unsigned int size;
            unsigned char packet[PCAP_MAX_PACKET_SIZE];

            /* handle received packets */
            if ((size = cls->pcap.ReadPacket(NULL, NULL, &proto, packet, PCAP_MAX_PACKET_SIZE)) >= sizeof(TIPv4ARP))
            {
               if (proto == 0x0806) /* ARP */
               {
                  unsigned char arp_src_mac[6];
                  IN_ADDR arp_src_ip;
                  IN_ADDR arp_dst_ip;

                  /* if ARP response is valid and parsed */
                  if (ARP_ParseResponse((TIPv4ARP *)packet, arp_src_mac, &arp_src_ip, NULL, &arp_dst_ip))
                  {
                     /* verify that the response is sent to our IP address */
                     if (arp_dst_ip.S_un.S_addr == adapter_ip.S_un.S_addr && arp_src_ip.S_un.S_addr == ip.S_un.S_addr)
                     {
                        response_time = GetTickCount() + MAX_ARP_SCAN_TIMEOUT - timeout;
                     }
                  }
               }
            }
         }

         /* close adapter */
         cls->pcap.CloseAdapter();
      }
      else
      {
         /* error opening adapter */
         __raise cls->OnScanError(E_COULD_NOT_OPEN_ADAPTER);
      }
   }
   else
   {
      /* error loading WinPCAP */
      __raise cls->OnScanError(E_WINPCAP_NOT_FOUND);
   }

   cls->scan_thread = NULL;
   return response_time;
}