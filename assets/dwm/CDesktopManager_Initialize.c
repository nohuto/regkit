__int64 __fastcall CDesktopManager::Initialize(CDesktopManager *this, struct IUnknown *a2)
{
  HANDLE EventW; // r15
  int v4; // eax
  signed int v5; // ebx
  __int64 v6; // rax
  bool v7; // r14
  __int64 v8; // rax
  LSTATUS v9; // ebx
  CWindowList *v10; // rax
  CWindowList *v11; // rbx
  CWindowList *v12; // rax
  int v13; // r9d
  CContactManager *v14; // rax
  CContactManager *v15; // rax
  CTransitionVisualController *v16; // rax
  const struct std::nothrow_t *v17; // rdx
  CTransitionVisualController *v18; // r11
  CAnimationScheduler *v19; // rax
  const struct std::nothrow_t *v20; // rdx
  CAnimationScheduler *v21; // r10
  CAnimationClockCoordinator *v22; // rax
  CAnimationClockCoordinator *v23; // r8
  CDesktopManager *v24; // rax
  int Factory; // eax
  CBaseObject *v26; // rcx
  CIconicBitmapRegistry *v27; // rax
  CIconicBitmapRegistry *v28; // rax
  CImmersiveIconicBitmapRegistry *v29; // rax
  CImmersiveIconicBitmapRegistry *v30; // rax
  __int64 v31; // rcx
  const struct std::nothrow_t *v32; // rdx
  CProjectionBorderManager *v33; // rax
  CProjectionBorderManager *v34; // rax
  __int64 v35; // rax
  signed int LastError; // eax
  int v37; // r9d
  HANDLE Thread; // rax
  signed int v39; // eax
  signed int v40; // eax
  unsigned int phkResult; // [rsp+20h] [rbp-50h]
  unsigned int phkResulta; // [rsp+20h] [rbp-50h]
  DWORD Type; // [rsp+30h] [rbp-40h] BYREF
  HKEY hKey; // [rsp+38h] [rbp-38h] BYREF
  __int64 v46; // [rsp+40h] [rbp-30h] BYREF
  __int64 pvParam; // [rsp+48h] [rbp-28h] BYREF
  HANDLE Handles[4]; // [rsp+50h] [rbp-20h] BYREF
  int v49; // [rsp+B8h] [rbp+48h] BYREF
  struct _RTL_CRITICAL_SECTION *Data; // [rsp+C0h] [rbp+50h] BYREF
  DWORD cbData; // [rsp+C8h] [rbp+58h] BYREF

  v46 = 0LL;
  EventW = 0LL;
  hKey = 0LL;
  v4 = ((__int64 (__fastcall *)(struct IUnknown *, GUID *, __int64 *))a2->lpVtbl->QueryInterface)(
         a2,
         &GUID_3ae5dff1_7681_484a_956a_6fd06c8e671e,
         &v46);
  v5 = v4;
  if ( v4 < 0 )
  {
    phkResult = 315;
    goto LABEL_108;
  }
  *((_BYTE *)this + 20) = 0;
  *((_BYTE *)this + 25) = 0;
  v6 = (*(__int64 (__fastcall **)(__int64))(*(_QWORD *)v46 + 24LL))(v46);
  *((_QWORD *)this + 9) = v6;
  v49 = 0;
  if ( (*(int (__fastcall **)(__int64, const wchar_t *, int *))(*(_QWORD *)v6 + 8LL))(
         v6,
         L"DefaultRemoteAppCreation",
         &v49) >= 0 )
    CDesktopManager::s_defaultRemoteAppCreation = v49 != 0;
  v7 = 0;
  v8 = wil::out_param<wil::unique_any_t<wil::details::unique_storage<wil::details::resource_policy<HKEY__ *,long (*)(HKEY__ *),&long RegCloseKey(HKEY__ *),wistd::integral_constant<unsigned __int64,0>,HKEY__ *,HKEY__ *,0,std::nullptr_t>>>>(
         Handles,
         &hKey);
  v9 = RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"Software\\Microsoft\\Windows\\Dwm", 0, 0x20019u, (PHKEY)(v8 + 8));
  wil::details::out_param_t<wil::unique_any_t<wil::details::unique_storage<wil::details::resource_policy<HKEY__ *,long (*)(HKEY__ *),&long RegCloseKey(HKEY__ *),wistd::integral_constant<unsigned __int64,0>,HKEY__ *,HKEY__ *,0,std::nullptr_t>>>>::~out_param_t<wil::unique_any_t<wil::details::unique_storage<wil::details::resource_policy<HKEY__ *,long (*)(HKEY__ *),&long RegCloseKey(HKEY__ *),wistd::integral_constant<unsigned __int64,0>,HKEY__ *,HKEY__ *,0,std::nullptr_t>>>>(Handles);
  if ( !v9 )
  {
    LODWORD(Data) = 0;
    cbData = 4;
    Type = 0;
    if ( !RegQueryValueExW(hKey, L"ForceEffectMode", 0LL, &Type, (LPBYTE)&Data, &cbData) && (_DWORD)Data == 2 )
      *((_BYTE *)this + 29) = 1;
    LODWORD(Data) = 0;
    cbData = 4;
    Type = 0;
    if ( !RegQueryValueExW(hKey, L"ForceUDwmSoftwareDevice", 0LL, &Type, (LPBYTE)&Data, &cbData) && (_DWORD)Data )
      v7 = 1;
    LODWORD(Data) = 0;
    cbData = 4;
    Type = 0;
    if ( !RegQueryValueExW(hKey, L"ForceDisableModeChangeAnimation", 0LL, &Type, (LPBYTE)&Data, &cbData) && (_DWORD)Data )
      CDesktopManager::s_forceDisableModeChangeAnimation = 1;
  }
  v10 = (CWindowList *)DefaultHeap::AllocClear(0x2C0uLL);
  v11 = v10;
  Data = (struct _RTL_CRITICAL_SECTION *)v10;
  if ( v10 )
  {
    memset_0(v10, 0, 0x2C0uLL);
    v12 = CWindowList::CWindowList(v11);
  }
  else
  {
    v12 = 0LL;
  }
  *((_QWORD *)this + 53) = v12;
  if ( v12 )
  {
    v14 = (CContactManager *)DefaultHeap::AllocClear(0x148uLL);
    Data = (struct _RTL_CRITICAL_SECTION *)v14;
    if ( v14 )
      v15 = CContactManager::CContactManager(v14);
    else
      v15 = 0LL;
    *((_QWORD *)this + 20) = v15;
    if ( !v15 )
    {
      phkResult = 396;
      goto LABEL_19;
    }
    v16 = (CTransitionVisualController *)DefaultHeap::AllocClear(0xC0uLL);
    Data = (struct _RTL_CRITICAL_SECTION *)v16;
    if ( v16 )
      v18 = CTransitionVisualController::CTransitionVisualController(v16);
    else
      v18 = 0LL;
    *((_QWORD *)CDesktopManager::s_pDesktopManagerInstance + 24) = v18;
    if ( !v18 )
    {
      phkResult = 399;
      goto LABEL_19;
    }
    v19 = (CAnimationScheduler *)operator new[](0x58uLL, v17);
    Data = (struct _RTL_CRITICAL_SECTION *)v19;
    if ( v19 )
      v21 = CAnimationScheduler::CAnimationScheduler(v19);
    else
      v21 = 0LL;
    *((_QWORD *)CDesktopManager::s_pDesktopManagerInstance + 23) = v21;
    if ( !v21 )
    {
      phkResult = 402;
      goto LABEL_19;
    }
    v22 = (CAnimationClockCoordinator *)operator new[](0x58uLL, v20);
    Data = (struct _RTL_CRITICAL_SECTION *)v22;
    if ( v22 )
      v23 = CAnimationClockCoordinator::CAnimationClockCoordinator(v22);
    else
      v23 = 0LL;
    v24 = CDesktopManager::s_pDesktopManagerInstance;
    *((_QWORD *)CDesktopManager::s_pDesktopManagerInstance + 21) = v23;
    if ( !v23 )
    {
      phkResult = 405;
      goto LABEL_19;
    }
    v4 = CAnimationClockCoordinator::SetEventCallback(
           v23,
           (struct IAnimationClockEventListener *)((*((_QWORD *)v24 + 23) + 8LL) & ((unsigned __int128)-(__int128)*((unsigned __int64 *)v24 + 23) >> 64)));
    v5 = v4;
    if ( v4 < 0 )
    {
      phkResult = 408;
    }
    else
    {
      pvParam = 8LL;
      if ( SystemParametersInfoW(0x48u, 8u, &pvParam, 0) )
        CDesktopManager::SetWindowAnimation(HIDWORD(pvParam) != 0);
      v4 = DwmRedirectionManagerInitialize(
             *((struct IDwmRedirectionClient **)this + 53),
             (struct IDwmRedirectionManager **)this + 8);
      v5 = v4;
      if ( v4 >= 0 )
      {
        wil::com_ptr_t<IDCompositionAnimationStats,wil::err_returncode_policy>::reset((char *)this + 48);
        Factory = CCompositor::Create((struct CCompositor **)this + 6);
        v5 = Factory;
        if ( Factory < 0 )
        {
          phkResulta = 426;
        }
        else
        {
          v26 = (CBaseObject *)*((_QWORD *)this + 7);
          *((_QWORD *)this + 7) = 0LL;
          if ( v26 )
            CBaseObject::Release(v26);
          Factory = CGraphicsDeviceManager::Create(v7, (struct CGraphicsDeviceManager **)this + 7);
          v5 = Factory;
          if ( Factory < 0 )
          {
            phkResulta = 429;
          }
          else
          {
            v27 = (CIconicBitmapRegistry *)DefaultHeap::AllocClear(0x70uLL);
            Data = (struct _RTL_CRITICAL_SECTION *)v27;
            if ( v27 )
              v28 = CIconicBitmapRegistry::CIconicBitmapRegistry(v27);
            else
              v28 = 0LL;
            *((_QWORD *)this + 28) = v28;
            if ( !v28 )
            {
              v5 = -2147024882;
              MilInstrumentationCheckHR_MaybeFailFast(0x14u, &dword_1800FD698, 1u, -2147024882, 0x1B5u, 0LL);
              goto LABEL_103;
            }
            Factory = CIconicBitmapRegistry::Init(v28);
            v5 = Factory;
            if ( Factory < 0 )
            {
              phkResulta = 438;
            }
            else
            {
              v29 = (CImmersiveIconicBitmapRegistry *)DefaultHeap::AllocClear(0x58uLL);
              Data = (struct _RTL_CRITICAL_SECTION *)v29;
              if ( v29 )
                v30 = CImmersiveIconicBitmapRegistry::CImmersiveIconicBitmapRegistry(v29);
              else
                v30 = 0LL;
              *((_QWORD *)this + 29) = v30;
              if ( !v30 )
              {
                v5 = -2147024882;
                MilInstrumentationCheckHR_MaybeFailFast(0x14u, &dword_1800FD698, 1u, -2147024882, 0x1B9u, 0LL);
                goto LABEL_103;
              }
              Factory = CImmersiveIconicBitmapRegistry::Init(v30);
              v5 = Factory;
              if ( Factory < 0 )
              {
                phkResulta = 442;
              }
              else
              {
                CDesktopManager::SetupDPIValues(this);
                *((_DWORD *)this + 118) = -1;
                if ( (Microsoft_Windows_Dwm_UdwmEnableBits & 1) != 0 )
                  McTemplateU0q_EtwEventWriteTransfer(v31, &UdwmStartup_Info, 1LL);
                CDesktopManager::UpdateRemotingMode(this);
                Factory = WICCreateImagingFactory_Proxy(567LL, (char *)this + 240);
                v5 = Factory;
                if ( Factory < 0 )
                {
                  phkResulta = 453;
                }
                else
                {
                  Factory = DWriteCreateFactory(0LL, &GUID_b859ee5a_d838_4b5b_a2e8_1adc7d93db48, (char *)this + 248);
                  v5 = Factory;
                  if ( Factory < 0 )
                  {
                    phkResulta = 460;
                  }
                  else
                  {
                    LODWORD(Data) = 13;
                    Factory = CDesktopManager::UpdateSettings(this, (unsigned int *)&Data);
                    v5 = Factory;
                    if ( Factory < 0 )
                    {
                      phkResulta = 475;
                    }
                    else
                    {
                      Factory = CLivePreview::Create((struct CLivePreview **)this + 57);
                      v5 = Factory;
                      if ( Factory < 0 )
                      {
                        phkResulta = 477;
                      }
                      else
                      {
                        v33 = (CProjectionBorderManager *)operator new[](0x290uLL, v32);
                        Data = (struct _RTL_CRITICAL_SECTION *)v33;
                        if ( v33 )
                          v34 = CProjectionBorderManager::CProjectionBorderManager(v33);
                        else
                          v34 = 0LL;
                        *((_QWORD *)this + 58) = v34;
                        if ( !*((_QWORD *)CDesktopManager::s_pDesktopManagerInstance + 58) )
                        {
                          v5 = -2147024882;
                          MilInstrumentationCheckHR_MaybeFailFast(0x14u, &dword_1800FD698, 1u, -2147024882, 0x1E0u, 0LL);
                          goto LABEL_103;
                        }
                        v35 = wil::out_param<wil::unique_any_t<wil::details::unique_storage<wil::details::resource_policy<HKEY__ *,long (*)(HKEY__ *),&long RegCloseKey(HKEY__ *),wistd::integral_constant<unsigned __int64,0>,HKEY__ *,HKEY__ *,0,std::nullptr_t>>>>(
                                Handles,
                                (char *)this + 208);
                        v5 = CCompositionEffectCache::Create((struct CCompositionEffectCache **)(v35 + 8));
                        wil::details::out_param_t<std::unique_ptr<CCompositionEffectCache>>::~out_param_t<std::unique_ptr<CCompositionEffectCache>>(Handles);
                        if ( v5 < 0 )
                        {
                          phkResulta = 482;
LABEL_79:
                          v37 = v5;
LABEL_102:
                          MilInstrumentationCheckHR_MaybeFailFast(0x14u, &dword_1800FD698, 1u, v37, phkResulta, 0LL);
LABEL_103:
                          CDesktopManager::NotifyRedirectionShutdown(this);
                          Data = &CDesktopManager::s_csDwmInstance;
                          LeaveCriticalSection(&CDesktopManager::s_csDwmInstance);
                          DwmRedirectionManagerShutdown();
                          EnterCriticalSection(&CDesktopManager::s_csDwmInstance);
                          if ( !EventW )
                            goto LABEL_110;
                          goto LABEL_104;
                        }
                        Factory = CWindowList::Initialize(*((CWindowList **)this + 53));
                        v5 = Factory;
                        if ( Factory < 0 )
                        {
                          phkResulta = 484;
                        }
                        else
                        {
                          SetLastError(0);
                          EventW = CreateEventW(0LL, 1, 0, 0LL);
                          if ( !EventW )
                          {
                            LastError = GetLastError();
                            v5 = LastError;
                            if ( LastError > 0 )
                              v5 = (unsigned __int16)LastError | 0x80070000;
                            phkResulta = 497;
LABEL_77:
                            if ( v5 >= 0 )
                              v5 = -2003304445;
                            goto LABEL_79;
                          }
                          SetLastError(0);
                          Thread = CreateThread(
                                     0LL,
                                     0LL,
                                     CDesktopManager::DwmEventThreadProc,
                                     EventW,
                                     0,
                                     (LPDWORD)this + 280);
                          *((_QWORD *)this + 141) = Thread;
                          if ( !Thread )
                          {
                            v39 = GetLastError();
                            v5 = v39;
                            if ( v39 > 0 )
                              v5 = (unsigned __int16)v39 | 0x80070000;
                            phkResulta = 506;
                            goto LABEL_77;
                          }
                          SetThreadDescription(Thread, L"uDWM Event Thread");
                          Handles[0] = EventW;
                          Handles[1] = *((HANDLE *)this + 141);
                          SetLastError(0);
                          if ( WaitForMultipleObjects(2u, Handles, 0, 0xFFFFFFFF) )
                          {
                            v40 = GetLastError();
                            v5 = v40;
                            if ( v40 > 0 )
                              v5 = (unsigned __int16)v40 | 0x80070000;
                            phkResulta = 523;
                            goto LABEL_77;
                          }
                          Factory = CDesktopManager::_InitializeWnf(this);
                          v5 = Factory;
                          if ( Factory >= 0 )
                          {
                            CDesktopManager::ReadProductType(this);
                            *((_BYTE *)this + 25) = 1;
LABEL_104:
                            CloseHandle(EventW);
                            goto LABEL_110;
                          }
                          phkResulta = 528;
                        }
                      }
                    }
                  }
                }
              }
            }
          }
        }
        v37 = Factory;
        goto LABEL_102;
      }
      phkResult = 422;
    }
LABEL_108:
    v13 = v4;
    goto LABEL_109;
  }
  phkResult = 393;
LABEL_19:
  v13 = -2147024882;
  v5 = -2147024882;
LABEL_109:
  MilInstrumentationCheckHR_MaybeFailFast(0x14u, &dword_1800FD698, 1u, v13, phkResult, 0LL);
LABEL_110:
  if ( v46 )
  {
    (*(void (__fastcall **)(__int64))(*(_QWORD *)v46 + 16LL))(v46);
    v46 = 0LL;
  }
  wil::details::unique_storage<wil::details::resource_policy<HKEY__ *,long (*)(HKEY__ *),&long RegCloseKey(HKEY__ *),wistd::integral_constant<unsigned __int64,0>,HKEY__ *,HKEY__ *,0,std::nullptr_t>>::~unique_storage<wil::details::resource_policy<HKEY__ *,long (*)(HKEY__ *),&long RegCloseKey(HKEY__ *),wistd::integral_constant<unsigned __int64,0>,HKEY__ *,HKEY__ *,0,std::nullptr_t>>(&hKey);
  return (unsigned int)v5;
}
