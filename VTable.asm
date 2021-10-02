	; ********************************************************************************************
	; Thunk Logic Summary. Four IPicture virtual functions will be subclassed:
	; IUnknown:Release - simply tracks ref count until zero. Once zero, it passes a message to the
	;	management window assigned during thunk creation. That window will forward a message
	;	to its client(s) to inform them that the passed IPicture is set to Nothing. The message is
	;	specifically formatted as: uMsg=WM_CancelMode: wParam=IPicture: lParam=uMsg
	;	This function is subclassed for both the IPicture & IPictureDisp interfaces
	; IPicture:get_Attributes - forces the function to identify the IPicture having transparency.
	; IPicture:Render - will draw the IPicture to the passed hDC. This thunk is only passed two types
	;	of IPicture object images: icon or 32bit premultiplied bitmaps. No other image type is passed.
	;	If icon, then DrawIconEx API used to render, else AlphaBlend API used
	;
	; This thunk lives in the project's thread until thread dies. 
	; If all its IPicture objects have been released, thunk will remain until the thread terminates
	; ********************************************************************************************
	[bits 32]
	; constants
		WM_CancelMode		equ dword 1Fh
		PICTURE_TRANSPARENT	equ dword 2h
		DI_NORMAL			equ dword 3h
		E_INVALIDARG		equ dword 80070057h

	; call stack			[ebp + 76]		pRcWBounds not used in thunk; therefore, not defined here
		%define pSrcHeight	[ebp + 72]		; cySrc parameter to IPicture:Render (himetric)
		%define pSrcWidth		[ebp + 68]		; cxSrc parameter to IPicture:Render (himetric)
		%define pSrcY		[ebp + 64]		; ySrc parameter to IPicture:Render (himetric)
		%define pSrcX		[ebp + 60]		; xSrc parameter to IPicture:Render (himetric)
		%define pCy			[ebp + 56]		; Cy parameter to IPicture:Render (pixel)
		%define pCx			[ebp + 52]		; Cx parameter to IPicture:Render (pixel)
		%define pY			[ebp + 48]		; Y parameter to IPicture:Render (pixel)
		%define pX			[ebp + 44]		; X parameter to IPicture:Render (pixel)
		%define pHDC		[ebp + 40]		; hDC parameter to IPicture:Render
		%define pAttrs		[ebp + 40]		; pAttrs parameter to IPicture:get_Attributes
		%define pUnk		[ebp + 36]		; IUnknown pointer
		%define lReturn		[ebp + 28]		; lReturn local, restored to eax after popad

	; storage - set during thunk creation
		%define hDC			[ebx +  0]		; memory DC
		%define addrIPicDisp	[ebx +  4]		; original IPictureDisp VTable this thunk is subclassing
		%define addrIPicture	[ebx +  8]		; original IPicture VTable this thunk is subclassing
		%define addrVtableCopy	[ebx + 12]		; memory location of VTable copy
		%define hWnd		[ebx + 16]		; management window
		%define vbDPI		[ebx + 20]		; equivalent to: 1440 / Screen.TwipsPerPixelX
		%define fnAlphaBlend	[ebx + 24]		; AlphaBlend function pointer
		%define fnPostMessage	[ebx + 28]		; PostMessage function pointer
		%define fnDrawIconEx	[ebx + 32]		; DrawIconEx function pointer
		%define fnSelectObject	[ebx + 36]		; SelectObject function pointer
		%define fnMulDiv		[ebx + 40]		; MulDiv function pointer
		%define tmpHandle		[ebx + 44]		; handle returned by IPicture:get_Handle
		%define tmpDC		[ebx + 48]		; DC being used (either memory or IPicture provided)
		%define tmpHeight		[ebx + 52]		; actual image height
		dd_hDC			dd 0
		dd_addrIPicDisp		dd 0
		dd_addrIPicture		dd 0
		dd_addrVtableCopy		dd 0
		dd_hWnd			dd 0
		dd_vbDPI			dd 0
		dd_fnAlphaBlend		dd 0
		dd_fnPostMessage		dd 0
		dd_fnDrawIconEx		dd 0
		dd_fnSelectObject		dd 0
		dd_fnMulDiv			dd 0
		dd_tmpHandle		dd 0
		dd_tmpDC			dd 0
		dd_tmpHeight		dd 0

	; Thunk start: IUnknown:Release for IPictureDisp (1 parameter)
		xor eax, eax					; setup the stack
		xor edx, edx
		pushad
		mov ebp, esp
		mov ebx, 012345678h
		xor esi, esi
		mov edi, addrIPicDisp
	_tstNothing:
		push dword pUnk					; call original IUnknown:Release, 8 bytes from VTable
		call [edi + 0x8]
		mov dword lReturn, eax
		cmp dword lReturn, esi
		jne _Rtn4
		push dword WM_CancelMode			; forward notification to management window
		push dword pUnk					; it will inform its client(s) that the object is now Nothing
		push dword WM_CancelMode
		push dword hWnd
		call fnPostMessage 
	_Rtn4:	
		popad
		ret 0x4

	Align 4
	; Thunk start: IUnknown:Release for IPicture (1 parameter)
		xor eax, eax					; setup the stack
		xor edx, edx
		pushad
		mov ebp, esp
		mov ebx, 012345678h
		xor esi, esi
		mov edi, addrIPicture
		jmp _tstNothing

	Align 4
	; Thunk start: get_Attributes (2 parameters)		
		xor eax, eax					; setup the stack
		xor edx, edx
		pushad
		mov ebp, esp
		mov ebx, 012345678h
		xor esi, esi
		cmp dword pAttrs, esi				; ensure no null pointer was passed
		je _ErrorAttrs
		mov edi, pAttrs
		mov dword [edi], PICTURE_TRANSPARENT	; set this attribute
		jmp _Rtn8
	_ErrorAttrs:
		mov dword lReturn, E_INVALIDARG
	_Rtn8:
		popad
		ret 0x8

	Align 4
	; Thunk start: Render (11 parameters)		
		xor eax, eax					; setup the stack
		xor edx, edx
		pushad
		mov ebp, esp
		mov ebx, 012345678h
		xor esi, esi
		mov esi, addrIPicture
		lea dword edi, tmpHandle			; get the image handle
		push dword edi
		push dword pUnk					
		call [esi + 0xC]					; function address for Get_Handle, 12 bytes from VTable
		cmp dword tmpHandle, 0x0			; abort if null handle
		jz _Rtn44						
	
		lea dword edi, tmpHeight			; get image height
		push edi
		push dword pUnk							
		call [esi + 0x1C]					; function address for Get_Height, 28 bytes from VTable

		lea dword edi, tmpDC				; get image type (2 bytes returned only, 3 bits used)
		push edi
		push dword pUnk							
		call [esi + 0x14]					; function address for Get_Type, 20 bytes from VTable

		and tmpDC, dword 0x7
		cmp tmpDC, dword 0x3				; check for icon
		je _DoDrawIconEx
		cmp tmpDC, dword 0x1				; check for bitmap vs unexpected
		jne _Rtn44

		push edi
		push dword pUnk
		call [esi + 0x28]					; function address for Get_CurDC, 40 bytes from VTable
		xor esi, esi
		cmp dword tmpDC, esi
		jne _DoAlphaBlend

		mov dWord eax, hDC				; select object into hDC
		mov dword tmpDC, eax
		push dword tmpHandle				; call SelectObject()
		push dword tmpDC
		call fnSelectObject
		mov dword tmpHandle, eax			; cache previous object (if any)

	_DoAlphaBlend:
		push dword 1FF0000h				; blend function (<32bpp format will fail to render)
		xor eax,eax						; negative render height
		sub dword eax, pSrcHeight			; src height is always negative; change to positive
		call _convertHimetrics				; convert from himetric & push onto stack
		push dword eax

		mov dword eax, pSrcWidth			; src width
		call _convertHimetrics				; convert from himetric & push onto stack
		push dword eax

		mov dword eax, tmpHeight			; actual height
		mov dword ecx, pSrcY				; srcY (relative from bottom of image)
		sub eax, ecx					; set offset relative to top of image
		call _convertHimetrics				; convert from himetric & push onto stack
		push dword eax

		cmp dword pSrcX, esi				; srcX
		jne _convertX
		push dword pSrcX
		jmp _doDest
	_convertX:
		mov dword eax, pSrcX				; convert srcX from himetric
		call _convertHimetrics				; convert from himetric & push onto stack
		push dword eax

	_doDest:
		push dword tmpDC					; src hDC
		push dword pCy					; dest height
		push dword pCx					; dest width
		push dword pY					; dest Y
		push dword pX					; dest X
		push dword pHDC					; dest hDC
		call fnAlphaBlend

		mov dword eax, hDC
		cmp dword tmpDC, eax				; did we select it into hDC?
		jne _Rtn44
		push dword tmpHandle				; call SelectObject()
		push dword tmpDC
		call fnSelectObject
		jmp _Rtn44

	_convertHimetrics:
		push dword 9ECh					; hardcoded conversion value (2540&)
		push dword vbDPI					; could use GetDevCaps API, but should use VB's DPI
		push dword eax					; himetric value
		call fnMulDiv
		ret

	_DoDrawIconEx:
		xor esi, esi
		push dword DI_NORMAL				; diFlags
		push dword esi					; hbrBrush
		push dword esi					; iStepIfAniCur
		push dword pCy					; cyHeight
		push dword pCx					; cxWidth 
		push dword tmpHandle				; hIcon
		push dword pY					; dest Y
		push dword pX					; dest X
		push dword pHDC					; dest hDC
		call fnDrawIconEx
	_Rtn44:
		popad
		ret 0x2C
