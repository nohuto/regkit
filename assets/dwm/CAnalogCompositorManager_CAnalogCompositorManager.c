CAnalogCompositorManager *__fastcall CAnalogCompositorManager::CAnalogCompositorManager(CAnalogCompositorManager *this)
{
  int v1; // eax
  __int64 v2; // rdx
  __int64 v3; // rax
  int AttachWatcher; // eax
  __int64 v5; // rdx
  __int64 v6; // rax
  int DetachWatcher; // eax
  int v9[34]; // [rsp+20h] [rbp-88h] BYREF
  wil::details::in1diag3 *retaddr; // [rsp+A8h] [rbp+0h]
  CAnalogCompositorManager *v11; // [rsp+B0h] [rbp+8h] BYREF
  void *v12; // [rsp+B8h] [rbp+10h]

  v11 = this;
  v12 = &qword_18011ECD0;
  qword_18011ECD0 = 0LL;
  Windows::Mirage::HolographicDriverDetectedWatcher::HolographicDriverDetectedWatcher((Windows::Mirage::HolographicDriverDetectedWatcher *)&unk_18011ECD8);
  qword_18011ED00 = 0LL;
  qword_18011ED08 = 0LL;
  qword_18011ED10 = 0LL;
  byte_18011ED18 = 0;
  qword_18011ED20 = 0LL;
  qword_18011ED28 = 0LL;
  qword_18011ED30 = 0LL;
  qword_18011ED38 = 0LL;
  qword_18011ED40 = 0LL;
  qword_18011ED48 = 0LL;
  qword_18011ED50 = 0LL;
  dword_18011ED58 = 0;
  byte_18011ED5C = 0;
  word_18011ED5D = 0;
  byte_18011ED5F = 0;
  qword_18011ED60 = 0LL;
  LODWORD(v11) = 0;
  if ( (*(int (__fastcall **)(_QWORD, const wchar_t *, CAnalogCompositorManager **))(**((_QWORD **)CDesktopManager::s_pDesktopManagerInstance
                                                                                      + 9)
                                                                                   + 8LL))(
         *((_QWORD *)CDesktopManager::s_pDesktopManagerInstance + 9),
         L"DisableHologramCompositor",
         &v11) < 0
    || !(_DWORD)v11 )
  {
    v1 = Windows::Mirage::HolographicDriverDetectedWatcher::RegisterForCMNotifications((Windows::Mirage::HolographicDriverDetectedWatcher *)&unk_18011ECD8);
    if ( v1 < 0 )
      wil::details::in1diag3::_FailFast_Hr(
        retaddr,
        (void *)0x26,
        (unsigned int)"clientcore\\windows\\dwm\\udwm\\analogcompositormanager.cpp",
        (const char *)(unsigned int)v1,
        v9[0]);
    LOBYTE(v2) = 0;
    v3 = wistd::function_void___cdecl_void__::function_void___cdecl_void____lambda_1fe009015b5481886de644cd00cd9360__void_(
           v9,
           v2);
    AttachWatcher = Windows::Mirage::HolographicDriverDetectedWatcher::CreateAttachWatcher(&unk_18011ECD8, v3);
    if ( AttachWatcher < 0 )
      wil::details::in1diag3::_FailFast_Hr(
        retaddr,
        (void *)0x2C,
        (unsigned int)"clientcore\\windows\\dwm\\udwm\\analogcompositormanager.cpp",
        (const char *)(unsigned int)AttachWatcher,
        v9[0]);
    LOBYTE(v5) = 0;
    v6 = wistd::function_void___cdecl_void__::function_void___cdecl_void____lambda_68ab246ca29dbf1f5c5163cf5c63f8ba__void_(
           v9,
           v5);
    DetachWatcher = Windows::Mirage::HolographicDriverDetectedWatcher::CreateDetachWatcher(&unk_18011ECD8, v6);
    if ( DetachWatcher < 0 )
      wil::details::in1diag3::_FailFast_Hr(
        retaddr,
        (void *)0x32,
        (unsigned int)"clientcore\\windows\\dwm\\udwm\\analogcompositormanager.cpp",
        (const char *)(unsigned int)DetachWatcher,
        v9[0]);
  }
  return (CAnalogCompositorManager *)&qword_18011ECD0;
}
