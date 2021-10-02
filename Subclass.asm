; ********************************************************************************************
; Thunk Logic Summary
;	A window is created and placed on the upper-most, top-level window in the process. This is to 
;	ensure the window stays alive until the process terminates. This window does nothing other than
;	listens for a message from the IPicture vTable thunk and, via a timer, checks to see if its VB client
;	is still alive. When the process terminates, this window will be destroyed also. During the 
;	destruction, it will clean up memory allocated for this window and the vTable thunk:
;
;	WM_CancelMode when received is a flag that a IPicture instance is now Nothing. This thunk will
;		call its client(s) and pass the notification for the client's own use/purpose.
;	WM_Timer when received will simply test to see if the client is still alive. If it is no longer alive, 
;		then the timer is destroyed.
;	WM_NCDestroy when received will be the last message before the Window is destroyed. During
;		destruction, the following actions are taken:
;		:: vTable thunk is released from memory
;		:: memory DC used by the thunk is destroyed
;		:: any GDI+ image handles that still exist are disposed. Array containing handles is destroyed
;		:: GDI+ is gracefully shut down
;		:: Final reference to GDI+ dll is released
; ********************************************************************************************
[bits 32]

	; constants
	%define WM_CancelMode	1Fh
	%define WM_Destroy	2h
	%define MEM_RELEASE	8000h
	%define GWL_WNDPROC     -4          ;SetWindowsLong WndProc parameter

	; call stack				
	%define pLParam		[ebp + 48]		; window message lParam
	%define pWParam		[ebp + 44]		; window message wParam
	%define pMsg		[ebp + 40]		; window message uMsg
	%define pHwnd		[ebp + 36]		; window message hWnd
	%define lReturn		[ebp + 28]		; lReturn local, restored to eax after popad

	; storage - set during thunk creation
	%define CBobject		[ebx +  0]		; client callback pointer
	%define CBfunction	[ebx +  4]		; client's function pointer
	%define addrVTableThunk	[ebx +  8]		; VTable thunk memory address
	%define addrTable		[ebx + 12]		; ptr to array of GDI+ handles
	%define fnEbMode 		[ebx + 16]		; IsBadCodePointer function address
	%define fnVirtualFree	[ebx + 20]		; VirtualFree function address
	%define fnFreeLibrary	[ebx + 24]		; FreeLibrary function address
	%define fnCallWndProc	[ebx + 28]		; CallWindowProc function address
	%define fnCoTaskMemFree	[ebx + 32]		; CoTaskMemFree function address
	%define fnDeleteDC	[ebx + 36]		; DeleteDC function address
	%define fnGDIpDispose	[ebx + 40]		; GDIpDisposeImage function address
	%define fnGDIpShutDown	[ebx + 44]		; GDIpShutDown function address
	%define hGDIpToken	[ebx + 48]		; GDI+ token
	%define hGDIpInst		[ebx + 52]		; GDI+ dll LoadLibrary instance
	%define hMsImgInst	[ebx + 56]
	dd_CBobject			dd 0
	dd_CBfunction		dd 0
	dd_addrVTableThunk	dd 0
	dd_addrTable		dd 0
	dd_fnfnEbMode 		dd 0
	dd_fnVirtualFree		dd 0
	dd_fnFreeLibrary		dd 0
	dd_fnCallWndProc		dd 0
	dd_fnCoTaskMemFree	dd 0
	dd_fnDeleteDC		dd 0
	dd_fnGDIpDispose		dd 0
	dd_fnGDIpShutDown		dd 0
	dd_hGDIpToken		dd 0
	dd_hGDIpInst		dd 0
	dd_hMsImgInst		dd 0

Align 4
	xor eax, eax				; setup stack
	xor edx, edx
	pushad
	mov ebp, esp
	mov ebx, 012345678h
	xor esi, esi
	nop
	push dword pLParam			; forward the message
	push dWord pWParam
	push dWord pMsg
	push dWord pHwnd
	push dWord 012345678h			; original WndProc patched from thunk creator
	call fnCallWndProc			; fnCallWndProc address patched from thunk creator
	mov dword lReturn, eax			; cache return value

	mov eax, WM_CancelMode
	cmp pMsg, eax				; WM_CancelMode
	jne _WMDestroy				; if not, check for WM_Destroy
	cmp dword pLParam, WM_CancelMode	; is it our formatted WM_CancelMode
	jne _Return					; if not, return

	cmp dword CBobject, esi			; client exists?
	je _Return
	cmp dword fnEbMode, esi			; not in this thunk if thunk compiled into exe
	je _notifyClient
	call fnEbMode           
    	cmp eax, dword 0x1           		; If EbMode = 1, running normally
    	je _notifyClient			
    	test eax, eax                		; If EbMode = 0, ended
    	jnz _Return					; else on a break point, skip notification
	mov dword CBobject, esi			; IDE terminated
	jmp _Return

_notifyClient:
	push dword pWParam			; push pUnk
	push dword CBobject			; push client reference
	call near CBfunction			; call client method
	jmp _Return

_WMDestroy:
	cmp pMsg, dword WM_Destroy		; WM_Destroy?
	jne _Return					; if not, return
	mov dword edi, addrVTableThunk	; get the memory DC used by the thunk (1st DWord at its address)
	mov dword eax, [edi]
	push eax
	call fnDeleteDC
	mov dword eax, [edi + 0xC]		; get location of VTable copy (4th DWord at address)
	push dword MEM_RELEASE			; release VTable copy
	push esi
	push eax
	call fnVirtualFree
	push dword MEM_RELEASE			; release VTable thunk
	push dword esi
	push dword addrVTableThunk
	call fnVirtualFree

	cmp dword addrTable, esi		; clear any GDI+ image handles, array of GDI+ handles exists?
	je _releaseToken				; if not, shut down GDI+
	mov edi, addrTable			; 1st DWord is active handle count, 3rd & later DWords are handles
	mov dword ecx, [edi]			; set array count
	cmp ecx, esi				; if count is zero, skip loop
	je _NextHandle
_loopHandles:
	mov dword eax, [edi + ecx * 0x4 + 0x4]	; get handle at next array slot
	cmp eax, esi				; if zero, skip disposal
	je _NextHandle			
	push ecx
	push dword eax				; dispose of the handle
	call fnGDIpDispose
	pop ecx
_NextHandle:
	loop _loopHandles				; continue looping until ecx=0
	push dword addrTable			; release the array
	call fnCoTaskMemFree

_releaseToken:
	cmp dword hGDIpToken, esi		; GDI+ running?
	je _releaseGDIp				; if not, release GDI+ DLL instance
	push dword hGDIpToken			; shut down GDI+
	call fnGDIpShutDown

_releaseGDIp:
	push dword hGDIpInst			; never non-zero
	call fnFreeLibrary			; release instance of GDI+ DLL
	push dword hMsImgInst			; never non-zero
	call fnFreeLibrary			; release instance of MSIMG32 DLL

_Return:
	popad
	Ret 0x10
