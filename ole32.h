#define FFI_LIB "ole32.dll"

typedef struct _GUID {
	uint32_t Data1;
	unsigned short Data2;
	unsigned short Data3;
	unsigned char Data4[8];
} GUID;

typedef GUID IID;
typedef GUID CLSID;

typedef IID *REFIID;
typedef CLSID *REFCLSID;

uint32_t CoInitializeEx(
	void *pvReserved,
	uint32_t dwCoInit
);
void CoUninitialize();

uint32_t CoCreateInstance(REFCLSID rclsid, void *pUnkOuter, uint32_t dwClsContext, REFIID riid, void **ppv);
uint32_t CoRegisterClassObject(REFCLSID rclsid, void *pUnk, uint32_t dwClsContext, uint32_t flags, uint32_t *lpdwRegister);
