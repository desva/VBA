#include "pch.h"
#include "Psapi.h"
#include "Shlwapi.h" // PathFindFileNameW 
#include <intrin.h>

//#pragma comment(linker,"/merge:.rdata=.data")
//#pragma comment(linker,"/merge:.text=.data")
#pragma comment(lib, "shlwapi.lib")


#define USERDTSC 0
#if USERDTSC==1 
#pragma intrinsic(__rdtsc)
#endif

#ifdef _M_X64 
	// 64 bit VBA
	// Found at RSP+b8(bytes!) = RSP+17 QWORDS (b8 bytes) [NOT or RSP-6 QWORDS 30h bytes]
#define PROCADDR UINT64
	typedef struct tPCT64 {
		UINT64* pParent;		// Points to parent
		ULONG32 filler1;		// 0
		ULONG32 nStack;			// size of stack - bytes?
		ULONG32 n1, n2;
	} vbaPCT;

	typedef struct tVbaFilename {
		WORD pad[0x13];			// Padding 
		wchar_t wszFilename;	// Filename
	} vbaFilename;

	typedef struct tObjTable64 {
		UINT64* pPtr1;
		vbaFilename* pOwner;	// gives spreadsheet file owner
		UINT64 padding[16];		// Padding
		char szProjName;		// VBAProject
		// remainder not relevant
	} vbaObjTable;

	typedef struct tParent64 {
		WORD nType;			// 2, 4 - flags?
		WORD wRefCount;		// Should equal 1
		ULONG32 wIndex;		// Should equal zero?
		vbaObjTable* pObjTbl;	// to Object table + 80 bytes / 10 qwords to get null terminated project name
		UINT64* pIdeData;
		UINT64* pPrvData;
		UINT64 marker;		// = FFFFFFFF FFFFFFFF (=-1)
		UINT64 marker2;		// = 0
		UINT64* pListEntry;	// pointer to entry in component collection - vbaCompEntry
		UINT64* pProjData;
		ULONG32 nProcs;		// number of procedures
		ULONG32 nProcs2;	// 0
		vbaPCT** procMap;	// points to pcode trailers. Null => imported DLL function
	} vbaParent;

	typedef struct tCompEntry64 {
		UINT64* pParent;	// points back to vbaParent
		UINT64 marker;		// = FFFFFFFF FFFFFFFF (=-1)
		UINT64* p1;			// ? Imports ?
		UINT64 flags;		// 0 ?
		UINT64* p2;			// ?
		UINT64 flags2;
		UINT64* pszModName;	// Null terminated module name
		DWORD nNumProcs;	// Number of procs in module
		DWORD nPadding;		// zero
		UINT64** ppszFnNames; // pointer to array of nNumProcs names
	} vbaCompEntry;

#else
	// 32-bit VBA
#define PROCADDR UINT32
	typedef struct tPCT32 {
		void* pParent;		// Back pointer to parent
		INT16 something;
		INT16 stacksize;
		INT16 nCount;		// Count of pcode bytes
		INT16 something3;
		INT16 somethingtest;
	} vbaPCT;

	typedef struct tVbaFilename {
		CHAR pad[0x16];		// Padding
		wchar_t wszFilename;	// Filename
	} vbaFilename;

	typedef struct tObjTable32 {
		void* pPtr1;
		void* pOwner;		// gives spreadsheet file owner
		char pad[0x38];		// Padding 
		char* pszProjName;	// VBAProject
		// remainder not relevant
	} vbaObjTable;

	typedef struct tParent32 {
		WORD wRefCount;		// Should equal 1
		WORD wIndex;
		void* pObjTbl;		// to Object table
		void* pIdeData;
		void* pPrvData;
		DWORD marker;		// -1
		DWORD marker2;		// 0
		void* pListEntry;	// pointer to entry in module collection
		void* pProjData;
		WORD nProcs;		// number of procedures
		WORD nProcs2;		// 0
		PCT** procMap;		// points to pcode trailers. Null => imported DLL function
		// Not used:
		DWORD nConsts;		// Number of constants
		DWORD nConstsToAlloc;
		void* pIdeData2;
		void* pIdeData3;
		void* pConsts;
	} vbaParent;

	typedef struct tCompEntry32 {
		void* pComp;		// points back to component, equals marker if component invalid.
		DWORD marker;		// -1/0xffffffff
		void* p1;			// ? Imports ?
		DWORD flags;		// 0 or all ff (true / false?)
		void* p2;			//?
		DWORD flags2;
		char* pszModName;	// Null terminated module name
		DWORD nNumProcs;	// Number of procs in module
		void** ppszFnNames; // pointer to array of nProcs names
	} vbaCompEntry;


#endif

typedef struct sVbaContext {
	vbaPCT* pct;
	vbaParent* pnt; // can be infered from pct
	vbaCompEntry* cpt;
} vbaContext;

// Timer
UINT64 ticks0;
UINT64 ticks;
UINT64 freq; // Adjusted so as to return time in microseconds

void InitTimer(){
	if (!freq) {
		QueryPerformanceFrequency((LARGE_INTEGER*)&freq);
		freq /= 1000000;
	}
}

UINT64 TimerMicroSeconds() { return (UINT64)(ticks / freq); }
UINT64 TimerToMicroSeconds(UINT64 ticks) { return (UINT64)(ticks / freq); }
UINT64 TimerToTicks(){ return ticks; }
#if USERDTSC==1 
void Start() { ticks0 = __rdstc(); }
void CTickTimer::Mark() {
	ticks = __rdtsc();
#else
void StartTimer() { QueryPerformanceCounter((LARGE_INTEGER*)&ticks0); }
void MarkTimer() {
	QueryPerformanceCounter((LARGE_INTEGER*)&ticks);
#endif
	UINT64 ticks1 = ticks;
	ticks -= ticks0;
	ticks0 = ticks1;
}

size_t nStatements = 0;		// Number of statements (BoS) executed
size_t nMethods = 0;		// Number of methods executed
vbaContext ctx;				// Context of currently executing VBA method
vbaPCT* prevPCT = 0;		// Previous P-code trailer. Use 0 to indicate we've just returned
wchar_t* prevPathFilename = 0; // cache this
wchar_t* prevFilename = 0;	// cache this
//map<UINT64, string> mapNames; // pct to fully-qualified name
//UINT64 measureIndex = 0;
//UINT64* pMeasures = 0;		// measure; timing
HANDLE g_hFile;


static void ClosePreviousContext(int IsReturn) {
	// Get here either via an explicit function exit (return, exception) 
	// or when execution has moved to a new function
	MarkTimer();
	//pMeasures[measureIndex++] = (UINT64)ctx.pct;
	//pMeasures[measureIndex++] = TimerToTicks();
	//measureIndex &= 0xffff;
	char szBuff[1536];
	//if (mapNames.end() == mapNames.find((UINT64)ctx.pct)) 
	{
		wchar_t* wszFilename = 0;
		if (prevPathFilename == &(ctx.pnt->pObjTbl->pOwner->wszFilename)) {
			wszFilename = prevFilename;
		}
		else {
			wszFilename = PathFindFileNameW(&(ctx.pnt->pObjTbl->pOwner->wszFilename));
			if (0 == wszFilename) {
				wszFilename = &(ctx.pnt->pObjTbl->pOwner->wszFilename);
			}
			prevFilename = wszFilename;
		}
		// linear scan :(
		DWORD i = 0, j = 0;
		while (i < ctx.cpt->nNumProcs)
		{
			if (ctx.pnt->procMap[i] == ctx.pct)
			{
				j = i;
				break;
			}
			i++;
		}
		wsprintfA(szBuff, "%s|[%S].%s:%s.%s|%d\n\0", (IsReturn==1 ? "R" : "E"),
			wszFilename, &(ctx.pnt->pObjTbl->szProjName), ctx.cpt->pszModName, ctx.cpt->ppszFnNames[j], TimerMicroSeconds());
		DWORD dwWritten;
#ifdef _DEBUG
		OutputDebugStringA(szBuff);
#else
		WriteFile(g_hFile, szBuff, lstrlenA(szBuff), &dwWritten, NULL);
#endif
	}
}

#ifdef _M_X64 
// 64-bit
static void ReturnHandler(UINT64 rsp) {
#else
// 32-bit - note different signature!!!
static void ReturnHandler() {
#endif
	prevPCT = 0;
	ClosePreviousContext(1);
}

#ifdef _M_X64 
// 64-bit
static void StatementHandler(UINT64 rsp) {
	rsp += 0xb8;
	vbaPCT* pct = (vbaPCT*)*(UINT64*)rsp;
#else
// 32-bit - note different signature!!!
static void StatementHandler() {
	DWORD pVBAStackFrame;
	__asm {
		mov pCodeAddr, esi;
		// On first opcode (and possibly subsequent) EBX points to code trailer, but ...
		//mov pCTAddr,ebx;			// ... EBX not guaranteed to be preserved across VBA method
		mov eax, DWORD PTR[ebp];	// Pushed previous EBP
		mov eax, DWORD PTR[eax];	// VBA stack frame - can walk this
		mov pVBAStackFrame, eax;
	}
	// 0x14(=050 bytes) on stack points to current P-Code Trailer (=EBX at function start). So too does -0x0c
	DWORD pPCTAddr = *((DWORD*)pVBAStackFrame - 0x14); // P-Code trail address
	vbaPCT* pct = (vbaPCT*)pPCTAddr; // ... and as a structure
#endif
	if (pct == prevPCT) return;	// Nothing to do if in same method
	if (0 != prevPCT) ClosePreviousContext(0);
	prevPCT = pct;
	ctx.pct = pct;
	ctx.pnt = (vbaParent*)pct->pParent;
	ctx.cpt = (vbaCompEntry*)ctx.pnt->pListEntry;
	StartTimer();
}

/*
// Trampoline for jump tables
// Manually assembled as x64 asm not supported
mov	qword ptr [rsp+8],rcx  ; prolog
push	rdi ; push registers we need to save
push 	rbx
push 	rbp
push 	rsi
push 	r12 ; vba stack pointer
push 	r13
push 	r14
push 	r15
mov		rcx, rsp
add 	rcx, 0x40 ; stack pointer adjusted for registers we've pushed, and index in to pct
sub		rsp, 0x20 ; reserved (shadow) space
movabs	rax, 0xeeee0000eeee0000 ; call our handler
call	rax
add	rsp, 0x20 ; remove shadow space
pop 	r15 ; restore
pop 	r14
pop 	r13
pop 	r12
pop 	rsi
pop 	rbp
pop 	rbx
pop	rdi
movabs 	rax, 0xffff0000ffff0000
jmp 	rax

0:  48 89 4c 24 08          mov    QWORD PTR [rsp+0x8],rcx
5:  57                      push   rdi
6:  53                      push   rbx
7:  55                      push   rbp
8:  56                      push   rsi
9:  41 54                   push   r12
b:  41 55                   push   r13
d:  41 56                   push   r14
f:  41 57                   push   r15
11: 48 89 e1                mov    rcx,rsp
14: 48 83 c1 40             add    rcx,0x40
18: 48 83 ec 20             sub    rsp,0x20
1c: 48 b8 00 00 ee ee 00    movabs rax,0xeeee0000eeee0000
23: 00 ee ee
26: ff d0                   call   rax
28: 48 83 c4 20             add    rsp,0x20
2c: 41 5f                   pop    r15
2e: 41 5e                   pop    r14
30: 41 5d                   pop    r13
32: 41 5c                   pop    r12
34: 5e                      pop    rsi
35: 5d                      pop    rbp
36: 5b                      pop    rbx
37: 5f                      pop    rdi
38: 48 b8 00 00 ff ff 00    movabs rax,0xffff0000ffff0000
3f: 00 ff ff
42: ff e0                   jmp    rax
*/

BYTE thunkFn64[0x47] =
{ 0x48, 0x89, 0x4C, 0x24, 0x08, 0x57, 0x53, 0x55, 0x56, 0x41, 0x54, 0x41, 0x55, 0x41, 0x56, 0x41,
	0x57, 0x48, 0x89, 0xE1, 0x48, 0x83, 0xC1, 0x40, 0x48, 0x83, 0xEC, 0x20, 0x48, 0xB8, 0x00, 0x00,
	0xEE, 0xEE, 0x00, 0x00, 0xEE, 0xEE, 0xFF, 0xD0, 0x48, 0x83, 0xC4, 0x20, 0x41, 0x5F, 0x41, 0x5E,
	0x41, 0x5D, 0x41, 0x5C, 0x5E, 0x5D, 0x5B, 0x5F, 0x48, 0xB8, 0x00, 0x00, 0xFF, 0xFF, 0x00, 0x00,
	0xFF, 0xFF, 0xFF, 0xE0 };

// 32-bit version
BYTE thunkFn32[0x12] = {
0x55,							// push ebp
0x8b, 0xec,						// mov ebp, esp
0x60,							// pusha
0x9c,							// pushf 9c
0xe8, 0x00, 0x00, 0x00, 0x00,	// call OutputOpcode (10)
0x9d,							// popf 9d
0x61,							// popa
0x5d,							// pop ebp
0xe9, 0x00, 0x00, 0x00, 0x00	// jmp pfn (20)
};

#ifdef _M_X64 
#define thunkFn thunkFn64
#define hndlraddr 0x1e
#define pcdaddr 0x3e
#else 
#define thunkFn thunkFn32
#define hndlraddr 0x06
#define pcdaddr 0x0e
#endif

BYTE*
	CreatePCodeThunk(PROCADDR pchdnler, int bIsReturnPCode) {
	BYTE* pCode = (BYTE*)VirtualAlloc(0, 0x100, MEM_COMMIT, PAGE_EXECUTE_READWRITE);
	if (pCode) {
		DWORD dwOldProtect;
		CopyMemory(pCode, thunkFn, sizeof(thunkFn));
		*(PROCADDR*)(pCode + hndlraddr) = (bIsReturnPCode == 1 ? (PROCADDR)ReturnHandler : (PROCADDR)StatementHandler)
#ifdef _M_X64 
			;
#else
			-(pCode + hndlraddr + 4);
#endif
		*(PROCADDR*)(pCode + pcdaddr) = pchdnler
#ifdef _M_X64 
			;
#else
			-(pCode + pcdaddr + 4);
#endif
		VirtualProtect(pCode, 0x100, PAGE_EXECUTE, &dwOldProtect);
		FlushInstructionCache(GetCurrentProcess(), 0, 0);
	}
	return pCode;
}

/*
	The p-code handler epilog looks like:
	lea rbx, tblDispatch ; This is a relative address
	jmp qword ptr [rbx+rax*8] ; rax holds the next opcode
	in binary, these look like (respectively):
	0x48, 0x8d, 0x1d, ll, hl, lh, hh ; last 4 bytes are offset relative to instruction pointer rip (points to first byte of opcode: 0x48)
	0xff, 0x24, 0xc3
	The BoS (Beginning of statement) opcode is 1338h bytes in
	and 3368h bytes in - an absolute address to jump to this time

	Exit / return Pcodes to patch
	lblEX_ExitForAryVar
	lblEX_ExitForCollObj
	lblEX_ExitForCollObj
	lblEX_ExitForCollObj
	lblEX_ExitForVar
	lblEX_ExitProcCb
	lblEX_ExitProcCbHresult
	lblEX_ExitProcCbStack
	lblEX_ExitProcCy
	lblEX_ExitProcCy
	lblEX_ExitProcCy
	lblEX_ExitProcCy
	lblEX_ExitProcCy
	lblEX_ExitProcFrameCb
	lblEX_ExitProcFrameCbHresult
	lblEX_ExitProcFrameCbStack
	lblEX_ExitProcHresult
	lblEX_ExitProcI2
	lblEX_ExitProcI4
	lblEX_ExitProcR4
	lblEX_ExitProcR8
	lblEX_ExitProcUI1
	lblEX_Return
	lblEX_ZeroRetVal
	lblEX_ZeroRetValVar

					case VBA_BoS:
				case VBA_ExitProcHResult:
				case VBA_ExitProcStr:
				case VBA_ExitProcI2:
				case VBA_ExitProcR4:
				case VBA_ExitProcR8:
				case VBA_ExitProcCy:
				case VBA_WordOpcodePref2:
				case VBA_WordOpcodePref3:
				case VBA_WordOpcodePref4:
				case VBA_WordOpcodePref5:

*/
// The corresponding **offsets** in bytes
// For the exit pcodes
__int64 PatchOffsets[0x19] = {
	0x2088, 0x2070, 0x2078, 0x2080, 0x2090, 0x1DC0, 0x3400, 0x1DC8,
	0x13A0, 0x13B0, 0x13B8, 0x13D0, 0x13D8, 0x2EB0, 0x3408, 0x2EB8,
	0xFC0, 0x1380, 0x1388, 0x1390, 0x1398, 0x1378, 0x1360, 0x33F0,
	0x33F8
};
// For the BoS pcodes
__int64 BosPatchOffsets[0x02] = {
	0x1338,0x3368
};
// Repeated opcodes
__int64 PcodeRepeats[0x13] = {
	1,3,1,1,1,1,5,1,1,1,1,1,1,1,1,1,1,1,1
};

UINT64 DispatchTable;
UINT64 DispatchTableSize; // in bytes! Divide by 8 to get number of entries
UINT64 CreatedThunks[0x13];
UINT64 ReplacedAddresses[0x13];
UINT64 BosThunk;
UINT64 BosOriginal;

int ActivateThunking() {
	UINT64 jmpTbl = DispatchTable;
	UINT64* pBoS1 = (UINT64*)(jmpTbl + BosPatchOffsets[0]);
	UINT64* pBoS2 = (UINT64*)(jmpTbl + BosPatchOffsets[1]);
	UINT64 pHndl = 0;
	DWORD dwOldProtect;
	VirtualProtect((BYTE*)jmpTbl, DispatchTableSize, PAGE_EXECUTE_READWRITE, &dwOldProtect);
	// Bos first
	BosOriginal = *(UINT64*)pBoS2;
	BosThunk = pHndl = (UINT64)CreatePCodeThunk(BosOriginal, 0);
#ifdef _DEBUG 
	char szBuff[1024];
	wsprintfA(szBuff, "VBAtrace BoS handler: 0x%p replaced with: 0x%p\n", pBoS1, &pHndl);
	OutputDebugStringA(szBuff);
#endif
	*pBoS1 = pHndl;
	*pBoS2 = pHndl;
	// Exits next
	int i = 0, j = 0, k = 0;
	while (i < 0x19) {
		k = 0;
		do {
			UINT64* pExit = (UINT64*)(jmpTbl + PatchOffsets[i]);
			if (k == 0) {
				pHndl = (UINT64)CreatePCodeThunk(*pExit, 1);
				CreatedThunks[j] = pHndl;
				ReplacedAddresses[j] = *pExit;
			}
			*pExit = pHndl;
			k++;
			i++;
		} while (k < PcodeRepeats[j]);
		j++;
	}
#ifdef _DEBUG 
	OutputDebugStringA("VBAtrace amended table.");
#endif
	FlushInstructionCache(GetCurrentProcess(), 0, 0);
	InitTimer();
	//pMeasures = (UINT64*)_aligned_malloc(0x10000, 32);
	return 1;
}

int Initialize() {
	// Create output trace if in release mode
	char szBuff[1024];
#ifndef _DEBUG
	char cTmpPath[] = "c:\\temp";	
	DWORD dwWritten;
	char szTempPath[MAX_PATH], szFilename[MAX_PATH];
	char* pszTemp = cTmpPath;
	if (0 != GetTempPathA(MAX_PATH, szTempPath))pszTemp = szTempPath;
	wsprintfA(szFilename, "%s\\VBATrace-%d.log", pszTemp, GetCurrentProcessId());
	g_hFile = CreateFileA(szFilename, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);
	if (INVALID_HANDLE_VALUE == g_hFile) return 0;
	wsprintfA(szBuff, "VBATrace (c) 2019-2021 Dr. D. Azzopardi\nAction(Exec of/Return from)|Method Name|Duration(µS)\n\0");
	WriteFile(g_hFile, szBuff, lstrlenA(szBuff), &dwWritten, NULL);
#endif
	// Find the VBE dll
	MODULEINFO modInfo;
	HMODULE hDLL = GetModuleHandleA("VBE7.DLL");

	if (NULL == hDLL) {
#ifdef _DEBUG 
		OutputDebugStringA("VBAtrace Unable to get VBA module handle!");
#endif
		return 0;
	}
	if (0 != GetModuleInformation(GetCurrentProcess(), hDLL, &modInfo, sizeof(modInfo))) {
		UINT64 nStart = (UINT64)modInfo.lpBaseOfDll;	// Base address
		UINT64* ptr = (UINT64*)(nStart);				// Pointer to base
		int nNumQwords = modInfo.SizeOfImage / 8;
		UINT64 nEnd = nStart + modInfo.SizeOfImage;
#ifdef _DEBUG 
		wsprintfA(szBuff, "VBAtrace VBE7.DLL loaded at 0x%x: 0x%x bytes", nStart, modInfo.SizeOfImage);
		OutputDebugStringA(szBuff);
#endif
		int nMaxSubseq = 0, nBest = 0;
		int iCurrRun = 0, iBestRun = 0;
		// 64-bit VBA uses a large jump table; each pcode is a 2-byte instruction
		// some are repeated
		// Searching for a table of > 1600 addresses that all point within VBE DLL
		// First create a map of candidates:
		typedef struct sTblRun {
			DWORD start;
			DWORD len;
		} TblRun;
		TblRun JmpTbls[256];
		int nCand = 0;
		int i;
		for (i = 0; i < nNumQwords; i++) {
			UINT64 nVal = *ptr++;
			if (nVal >= nStart && nVal <= nEnd) {
				if (nMaxSubseq == 0) {		// starting a new run
					iCurrRun = i;
				}
				nMaxSubseq++;
			}
			else {
				if (nMaxSubseq > 255) {
					JmpTbls[nCand].start = iCurrRun;
					JmpTbls[nCand].len = nMaxSubseq;
					nCand++;
					if (nCand > 255)return 0; // out of bounds, too many runs
				}
				if (nMaxSubseq > nBest) {
					nBest = nMaxSubseq;
					iBestRun = iCurrRun;
				}
				nMaxSubseq = 0;
			}
		}
		// Iterate through candidates
		for(i=0;i<nCand;i++) {
#ifdef _DEBUG 
			wsprintfA(szBuff, "VBAtrace Found %d entry table at 0x%x", JmpTbls[i].len,  nStart + (8 * JmpTbls[i].start));
			OutputDebugStringA(szBuff);
#endif
			UINT64 jmpTbl = (UINT64)(nStart + (8 * (UINT64)JmpTbls[i].start));
			UINT64* pBoS1 = (UINT64*)(jmpTbl + BosPatchOffsets[0]); // 0x1338
			UINT64* pBoS2 = (UINT64*)(jmpTbl + BosPatchOffsets[1]); // 0x3368
			if (*(UINT64*)pBoS2 == *(UINT64*)pBoS1) {
				DispatchTable = jmpTbl;
				DispatchTableSize = 8 * (UINT64)JmpTbls[i].len;
				return ActivateThunking();
			}
		}
	}
	return 0;
}

BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
	switch (ul_reason_for_call) {
	case DLL_PROCESS_ATTACH:
		//__security_init_cookie();
#ifdef _DEBUG 
		OutputDebugStringA("VBAtrace Debug DLL compiled : " __DATE__ "\n");
#else
		OutputDebugStringA("VBAtrace Release DLL compiled : " __DATE__ "\n");
#endif
		// Look for and patch jump tables
		Initialize();
		break;
	default:
	case DLL_THREAD_ATTACH:
#ifdef _DEBUG 
		OutputDebugStringA("VBAtrace thread attach\n");
#endif
		break;
	}
	return TRUE;	
}