#define  DLT_EN10MB        1        /* Ethernet (10Mb) */
#define  PCAP_ERRBUF_SIZE  256      /* Size to use when allocating the buffer that contains the libpcap errors */
#define  PCAP_OPENFLAG_PROMISCUOUS     1
#define  PCAP_MAX_PACKET_SIZE          (65536)

typedef struct pcap pcap_t; ///< Descriptor of an open capture instance. This structure is \b opaque to the user, that handles its content through the functions provided by wpcap.dll.
typedef struct pcap_dumper pcap_dumper_t; ///< libpcap savefile descriptor.
typedef struct pcap_if pcap_if_t; ///< Item in a list of interfaces, see pcap_if
typedef struct pcap_addr pcap_addr_t; ///< Representation of an interface address, see pcap_addr

#define PCAP_IF_LOOPBACK	0x00000001	///< interface is loopback

struct pcap_addr {
	struct pcap_addr *next;       // if not NULL, a pointer to the next element in the list; NULL for the last element of the list
	struct sockaddr *addr;		   // a pointer to a struct sockaddr containing an address
	struct sockaddr *netmask;	   // if not NULL, a pointer to a struct sockaddr that contains the netmask corresponding to the address pointed to by addr.
	struct sockaddr *broadaddr;	// if not NULL, a pointer to a struct sockaddr that contains the broadcast address corre­ sponding to the address pointed to by addr; may be null if the interface doesn't support broadcasts
	struct sockaddr *dstaddr;	   // if not NULL, a pointer to a struct sockaddr that contains the destination address corre­ sponding to the address pointed to by addr; may be null if the interface isn't a point- to-point interface
};

struct pcap_if
{
   struct pcap_if *next;         // if not NULL, a pointer to the next element in the list; NULL for the last element of the list
   char *name;		               // a pointer to a string giving a name for the device to pass to pcap_open_live()
   char *description;	         // if not NULL, a pointer to a string giving a human-readable description of the device
   struct pcap_addr *addresses;  // a pointer to the first element of a list of addresses for the interface
   u_int flags;		            // PCAP_IF_ interface flags. Currently the only possible flag is \b PCAP_IF_LOOPBACK, that is set if the interface is a loopback interface.
};

struct pcap_pkthdr
{
   struct timeval ts;  
   u_int caplen; 
   u_int len;    
};

#pragma pack(1)
struct  TEthernetPacket
{
   BYTE     h_dest[6];     /* destination eth addr	*/
   BYTE     h_source[6];   /* source ether addr	*/
   WORD     h_proto;       /* packet type ID field	*/
};
#pragma pack()

/* the class-wrapper around basic functionality of WinPCAP library */
class TWinPCAP
{
private:
   HMODULE        h_pcap_lib;
   pcap_if_t *    alldevs;
   pcap_t *       adapter;
   u_char         adapter_mac[6];   /* mac of opened adapter */
   char           errbuf[PCAP_ERRBUF_SIZE];

   /* winpcap exports */
   int (*pcap_findalldevs)(pcap_if_t **alldevsp, char *errbuf);
   void (*pcap_freealldevs)(pcap_if_t *alldevsp);
   int (*pcap_datalink)(pcap_t *p);
   char * (*pcap_geterr)(pcap_t *p);
   void (*pcap_close)(pcap_t *p);
   pcap_t * (*pcap_open)(const char * source, int snaplen, int flags, int read_timeout, struct pcap_rmtauth * auth, char * errbuf);
   int (*pcap_next_ex)(pcap_t * p, struct pcap_pkthdr ** pkt_header, const u_char ** pkt_data);
   int (*pcap_sendpacket)(pcap_t * p, const u_char * pkt_data, int pkt_size);

   /* helper functions */
   bool GetAllDevs(void);
   pcap_if_t * GetAdapter(int adapter);

public:
   TWinPCAP(void);
   ~TWinPCAP(void);

   bool IsLoaded(void) { return (this->h_pcap_lib != NULL); }
   bool LoadPCAP(void);
   int GetAdapterCount(void);
   char * GetAdapterName(int adapter);
   char * GetAdapterDescription(int adapter);
   bool GetAdapterMAC(int adapter, unsigned char * mac);
   bool GetAdapterIP(int adapter, IN_ADDR * ip, IN_ADDR * mask);
   bool OpenAdapter(int adapter);
   bool CloseAdapter(void);
   bool IsEthernet(void);
   bool SendPacket(unsigned char * src_mac, unsigned char * dst_mac, int proto, unsigned char * payload, unsigned int payload_size);
   unsigned int ReadPacket(unsigned char * src_mac, unsigned char * dst_mac, int * proto, unsigned char * payload , unsigned int buffer_size);
};

extern void MAC2TXT(char * text, unsigned char * mac);
extern bool TXT2MAC(unsigned char * mac, char * text);