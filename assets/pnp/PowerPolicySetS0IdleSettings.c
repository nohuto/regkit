// local variable allocation has failed, the output may be wrong!
int __fastcall FxPkgPnp::PowerPolicySetS0IdleSettings(FxPkgPnp *this, _WDF_DEVICE_POWER_POLICY_IDLE_SETTINGS *Settings)
{
  FxPowerPolicyOwnerSettings *m_Owner; // r14
  _WDF_TRI_STATE v3; // eax
  IdleTimeoutManagement *p_m_TimeoutMgmt; // r14
  _WDF_DEVICE_POWER_POLICY_IDLE_SETTINGS *v6; // rdi
  unsigned __int8 v7; // r12
  int DxState; // esi
  unsigned __int8 Set; // r15
  int result; // eax
  _DEVICE_POWER_STATE _a2; // eax
  unsigned int v12; // r8d
  signed int inited; // eax
  signed int v14; // r15d
  unsigned int IdleTimeout; // r13d
  _WDF_POWER_POLICY_S0_IDLE_USER_CONTROL UserControlOfIdleSettings; // eax
  unsigned __int8 m_DirectedTransitionsSupported; // al
  _FX_DRIVER_GLOBALS *m_Globals; // rax
  bool v19; // zf
  unsigned __int64 m_PoFxDeviceFlags; // rax
  unsigned __int64 v21; // rax
  _WDF_TRI_STATE PowerUpIdleDeviceOnSystemWake; // ecx
  _FX_DRIVER_GLOBALS *v23; // rcx
  unsigned __int16 v24; // r9
  const char *v25; // rcx
  FxPowerPolicyOwnerSettings *v26; // rcx
  _WDF_POWER_POLICY_S0_IDLE_CAPABILITIES IdleCaps; // eax
  __int64 i; // rcx
  unsigned int ExcludeD3Cold; // ecx
  void (__fastcall *SetD3ColdSupport)(void *, unsigned __int8); // rax
  const _GUID *traceGuid; // [rsp+20h] [rbp-E0h]
  int v32; // [rsp+38h] [rbp-C8h]
  unsigned __int8 enabled; // [rsp+40h] [rbp-C0h] BYREF
  unsigned __int8 directedTransitions; // [rsp+41h] [rbp-BFh] BYREF
  unsigned __int8 dfxChildrenOptional; // [rsp+42h] [rbp-BEh] BYREF
  unsigned __int8 v36; // [rsp+43h] [rbp-BDh]
  unsigned __int8 useWdfTimerForPofx; // [rsp+44h] [rbp-BCh] BYREF
  unsigned __int8 v38; // [rsp+45h] [rbp-BBh]
  _UNICODE_STRING valueName; // [rsp+48h] [rbp-B8h] BYREF
  _UNICODE_STRING childrenOptionalName; // [rsp+58h] [rbp-A8h] BYREF
  _UNICODE_STRING useWdfTimerForPofxName; // [rsp+68h] [rbp-98h] BYREF
  _BYTE useWdfTimerForPofxName_buffer[48]; // [rsp+78h] [rbp-88h] OVERLAPPED BYREF
  __int64 v43; // [rsp+A8h] [rbp-58h]
  wchar_t v44; // [rsp+B0h] [rbp-50h]
  wchar_t childrenOptionalName_buffer[48]; // [rsp+C0h] [rbp-40h] BYREF
  _OWORD v46[4]; // [rsp+120h] [rbp+20h] BYREF
  wchar_t v47; // [rsp+160h] [rbp+60h]

  m_Owner = this->m_PowerPolicyMachine.m_Owner;
  v3 = Settings->Enabled;
  v36 = 0;
  p_m_TimeoutMgmt = &m_Owner->m_IdleSettings.m_TimeoutMgmt;
  v6 = Settings;
  v7 = 0;
  DxState = 4;
  if ( v3 == WdfTrue )
  {
    enabled = 1;
  }
  else if ( v3 == WdfUseDefault )
  {
    enabled = 1;
    if ( KeGetCurrentIrql() )
    {
      WPP_IFR_SF_(this->m_Globals, 3u, 0xCu, 0x2Fu, (const _GUID *)&WPP_FxPkgPnp_cpp_Traceguids);
    }
    else
    {
      v44 = aWdfdefaultidle[28];
      *(_OWORD *)useWdfTimerForPofxName_buffer = *(_OWORD *)L"WdfDefaultIdleInWorkingState";
      *(_QWORD *)&valueName.Length = 3801144LL;
      *(_OWORD *)&useWdfTimerForPofxName_buffer[16] = *(_OWORD *)L"ltIdleInWorkingState";
      valueName.Buffer = (wchar_t *)useWdfTimerForPofxName_buffer;
      *(_OWORD *)&useWdfTimerForPofxName_buffer[32] = *(_OWORD *)L"WorkingState";
      v43 = *(_QWORD *)L"tate";
      FxPkgPnp::ReadRegistryS0Idle(this, &valueName, &enabled);
    }
  }
  else
  {
    enabled = 0;
  }
  Set = this->m_PowerPolicyMachine.m_Owner->m_IdleSettings.Set;
  v38 = Set;
  if ( !this->m_CapsQueried && !KeGetCurrentIrql() )
  {
    result = FxPkgPnp::QueryForCapabilities(this);
    if ( result < 0 )
      return result;
  }
  if ( v6->IdleCaps == IdleCannotWakeFromS0 )
  {
    DxState = v6->DxState;
    v36 = 0;
    if ( DxState == 5 )
      DxState = 4;
    goto LABEL_31;
  }
  if ( (unsigned int)(v6->IdleCaps - 2) > 1 )
    goto LABEL_31;
  DxState = v6->DxState;
  v36 = 1;
  _a2 = FxPkgPnp::PowerPolicyGetDeviceDeepestDeviceWakeState(this, PowerSystemWorking);
  if ( DxState == 5 )
  {
    DxState = _a2;
    if ( (unsigned int)(_a2 - 2) > 2 )
    {
LABEL_17:
      WPP_IFR_SF_DD(
        this->m_Globals,
        (unsigned __int8)Settings,
        0xCu,
        0x30u,
        (const _GUID *)&WPP_FxPkgPnp_cpp_Traceguids,
        _a2,
        -1073741101);
      return -1073741101;
    }
    if ( _a2 > PowerDeviceD2 )
    {
      if ( v6->IdleCaps == IdleUsbSelectiveSuspend )
        goto LABEL_17;
      goto LABEL_31;
    }
  }
  else
  {
    if ( DxState > _a2 )
    {
      WPP_IFR_SF_LLd(this->m_Globals, (unsigned __int8)Settings, v12, 0x31u, traceGuid, DxState, _a2, v32);
      return -1073741101;
    }
    if ( DxState > 3 )
    {
      if ( v6->IdleCaps == IdleUsbSelectiveSuspend )
      {
        WPP_IFR_SF_DD(
          this->m_Globals,
          (unsigned __int8)Settings,
          0xCu,
          0x32u,
          (const _GUID *)&WPP_FxPkgPnp_cpp_Traceguids,
          DxState,
          -1073741101);
        return -1073741101;
      }
      goto LABEL_31;
    }
  }
  if ( v6->IdleCaps == IdleUsbSelectiveSuspend )
  {
    inited = FxPowerPolicyMachine::InitUsbSS(&this->m_PowerPolicyMachine);
    v14 = inited;
    if ( inited < 0 )
    {
      WPP_IFR_SF_D(this->m_Globals, 2u, 0xCu, 0x33u, (const _GUID *)&WPP_FxPkgPnp_cpp_Traceguids, inited);
      return v14;
    }
    Set = v38;
  }
LABEL_31:
  IdleTimeout = v6->IdleTimeout;
  if ( !IdleTimeout )
    IdleTimeout = 5000;
  UserControlOfIdleSettings = v6->UserControlOfIdleSettings;
  if ( UserControlOfIdleSettings == IdleAllowUserControl )
  {
    result = FxPkgPnp::UpdateWmiInstanceForS0Idle(this, AddInstance);
    if ( result < 0 )
      return result;
    if ( v6->Enabled == WdfUseDefault )
    {
      if ( Set || KeGetCurrentIrql() )
      {
        enabled = this->m_PowerPolicyMachine.m_Owner->m_IdleSettings.Enabled;
      }
      else
      {
        *(_DWORD *)&useWdfTimerForPofxName_buffer[32] = *(_DWORD *)L"te";
        *(_WORD *)&useWdfTimerForPofxName_buffer[36] = aIdleinworkings[18];
        valueName.Buffer = (wchar_t *)useWdfTimerForPofxName_buffer;
        *(_OWORD *)useWdfTimerForPofxName_buffer = *(_OWORD *)L"IdleInWorkingState";
        *(_QWORD *)&valueName.Length = 2490404LL;
        *(_OWORD *)&useWdfTimerForPofxName_buffer[16] = *(_OWORD *)L"rkingState";
        FxPkgPnp::ReadRegistryS0Idle(this, &valueName, &enabled);
      }
    }
    v7 = 1;
  }
  else if ( UserControlOfIdleSettings == IdleDoNotAllowUserControl )
  {
    FxPkgPnp::UpdateWmiInstanceForS0Idle(this, RemoveInstance);
  }
  if ( !Set )
  {
    this->m_PowerPolicyMachine.m_Owner->m_IdleSettings.Set = 1;
    this->m_PowerPolicyMachine.m_Owner->m_IdleSettings.Overridable = v7;
  }
  if ( v6->Size <= 0x1C )
    goto LABEL_62;
  if ( !Set )
  {
    if ( (unsigned int)(v6->IdleTimeoutType - 1) <= 1 )
    {
      result = IdleTimeoutManagement::UseSystemManagedIdleTimeout(p_m_TimeoutMgmt, this->m_Globals);
      if ( result < 0 )
        return result;
      if ( (p_m_TimeoutMgmt->m_IdleTimeoutStatus & 4) != 0 )
      {
        m_DirectedTransitionsSupported = p_m_TimeoutMgmt->m_DirectedTransitionsSupported;
      }
      else
      {
        m_DirectedTransitionsSupported = unk_1C009FF62;
        if ( this->m_Globals->WdfBindInfo->Version.Minor >= 0x1F )
          m_DirectedTransitionsSupported = 1;
      }
      directedTransitions = m_DirectedTransitionsSupported;
      m_Globals = this->m_Globals;
      dfxChildrenOptional = 0;
      useWdfTimerForPofx = m_Globals->WdfBindInfo->Version.Minor >= 0x21;
      if ( (p_m_TimeoutMgmt->m_IdleTimeoutStatus & 4) != 0 )
        dfxChildrenOptional = (p_m_TimeoutMgmt->m_PoFxDeviceFlags & 6) == 6;
      if ( KeGetCurrentIrql() )
      {
        WPP_IFR_SF_(this->m_Globals, 3u, 0xCu, 0x34u, (const _GUID *)&WPP_FxPkgPnp_cpp_Traceguids);
      }
      else
      {
        v46[0] = *(_OWORD *)L"WdfDirectedPowerTransitionEnable";
        v46[1] = *(_OWORD *)L"tedPowerTransitionEnable";
        v47 = aWdfdirectedpow_0[32];
        v46[2] = *(_OWORD *)L"TransitionEnable";
        v46[3] = *(_OWORD *)L"onEnable";
        valueName.Buffer = (wchar_t *)v46;
        *(_QWORD *)&valueName.Length = 4325440LL;
        FxPkgPnp::ReadRegistryS0Idle(this, &valueName, &directedTransitions);
        wcscpy(childrenOptionalName_buffer, L"WdfDirectedPowerTransitionChildrenOptional");
        *(_QWORD *)&childrenOptionalName.Length = 5636180LL;
        childrenOptionalName.Buffer = childrenOptionalName_buffer;
        FxPkgPnp::ReadRegistryS0Idle(this, &childrenOptionalName, &dfxChildrenOptional);
        *(_DWORD *)&useWdfTimerForPofxName_buffer[40] = *(_DWORD *)L"x";
        *(_OWORD *)useWdfTimerForPofxName_buffer = *(_OWORD *)L"WdfUseWdfTimerForPofx";
        *(_QWORD *)&useWdfTimerForPofxName_buffer[32] = *(_QWORD *)L"rPofx";
        *(_OWORD *)&useWdfTimerForPofxName_buffer[16] = *(_OWORD *)L"fTimerForPofx";
        *(_QWORD *)&useWdfTimerForPofxName.Length = 2883626LL;
        useWdfTimerForPofxName.Buffer = (wchar_t *)useWdfTimerForPofxName_buffer;
        FxPkgPnp::ReadRegistryS0Idle(this, &useWdfTimerForPofxName, &useWdfTimerForPofx);
      }
      v19 = dfxChildrenOptional == 0;
      p_m_TimeoutMgmt->m_DirectedTransitionsSupported = directedTransitions;
      m_PoFxDeviceFlags = p_m_TimeoutMgmt->m_PoFxDeviceFlags;
      if ( v19 )
        v21 = m_PoFxDeviceFlags & 0xFFFFFFFFFFFFFFF9uLL;
      else
        v21 = m_PoFxDeviceFlags | 6;
      p_m_TimeoutMgmt->m_PoFxDeviceFlags = v21;
      Set = v38;
      p_m_TimeoutMgmt->m_UseWdfTimerForPofx = useWdfTimerForPofx;
    }
LABEL_62:
    if ( v6->IdleCaps != IdleCannotWakeFromS0 || v6->Size <= 0x18 )
      goto LABEL_75;
    PowerUpIdleDeviceOnSystemWake = v6->PowerUpIdleDeviceOnSystemWake;
    if ( PowerUpIdleDeviceOnSystemWake )
    {
      if ( PowerUpIdleDeviceOnSystemWake != WdfTrue )
        goto LABEL_75;
      this->m_PowerPolicyMachine.m_Owner->m_IdleSettings.PowerUpIdleDeviceOnSystemWake = 1;
      v23 = this->m_Globals;
      if ( !v23->FxVerboseOn )
        goto LABEL_75;
      v24 = 54;
    }
    else
    {
      this->m_PowerPolicyMachine.m_Owner->m_IdleSettings.PowerUpIdleDeviceOnSystemWake = 0;
      v23 = this->m_Globals;
      if ( !v23->FxVerboseOn )
      {
LABEL_75:
        v26 = this->m_PowerPolicyMachine.m_Owner;
        if ( !v26->m_IdleSettings.UsbSSCapabilityKnown )
        {
          IdleCaps = v6->IdleCaps;
          if ( IdleCaps == IdleUsbSelectiveSuspend )
          {
            for ( i = 0LL; i < 2; ++i )
              *(&this->m_PowerPolicyMachine.m_Owner->m_IdleSettings.UsbSSCapable + i) = 1;
          }
          else if ( IdleCaps == IdleCanWakeFromS0 )
          {
            v26->m_IdleSettings.UsbSSCapabilityKnown = 1;
          }
        }
        this->m_PowerPolicyMachine.m_Owner->m_IdleSettings.WakeFromS0Capable = v36;
        this->m_PowerPolicyMachine.m_Owner->m_IdleSettings.DxState = DxState;
        if ( (p_m_TimeoutMgmt->m_IdleTimeoutStatus & 2) == 0 || p_m_TimeoutMgmt->m_UseWdfTimerForPofx )
        {
          this->m_PowerPolicyMachine.m_Owner->m_PowerIdleMachine.m_PowerTimeout = (_LARGE_INTEGER)(-10000LL * IdleTimeout);
        }
        else
        {
          if ( !Set )
            this->m_PowerPolicyMachine.m_Owner->m_PowerIdleMachine.m_PowerTimeout = (_LARGE_INTEGER)-1LL;
          if ( v6->IdleTimeoutType == SystemManagedIdleTimeoutWithHint )
            this->m_PowerPolicyMachine.m_Owner->m_PoxInterface.m_NextIdleTimeoutHint = IdleTimeout;
        }
        if ( v6->Size > 0x1C )
        {
          ExcludeD3Cold = v6->ExcludeD3Cold;
          if ( ExcludeD3Cold != 2 )
          {
            if ( ExcludeD3Cold )
            {
              if ( ExcludeD3Cold != 1 )
                WPP_IFR_SF_D(
                  this->m_Globals,
                  2u,
                  0xCu,
                  0x38u,
                  (const _GUID *)&WPP_FxPkgPnp_cpp_Traceguids,
                  ExcludeD3Cold);
              LOBYTE(Settings) = 0;
            }
            else
            {
              LOBYTE(Settings) = 1;
            }
            this->m_PowerPolicyMachine.m_Owner->m_IdleSettings.D3ColdCapabilityKnown = 1;
            this->m_PowerPolicyMachine.m_Owner->m_IdleSettings.D3ColdSupported = (unsigned __int8)Settings;
            SetD3ColdSupport = this->m_D3ColdInterface.SetD3ColdSupport;
            if ( SetD3ColdSupport )
              ((void (__fastcall *)(void *, _WDF_DEVICE_POWER_POLICY_IDLE_SETTINGS *, __int64))SetD3ColdSupport)(
                this->m_D3ColdInterface.Context,
                Settings,
                1LL);
          }
        }
        FxPkgPnp::PowerPolicySetS0IdleState(this, enabled);
        return 0;
      }
      v24 = 55;
    }
    WPP_IFR_SF_(v23, 5u, 0xCu, v24, (const _GUID *)&WPP_FxPkgPnp_cpp_Traceguids);
    goto LABEL_75;
  }
  Settings = (_WDF_DEVICE_POWER_POLICY_IDLE_SETTINGS *)(p_m_TimeoutMgmt->m_IdleTimeoutStatus & 2);
  if ( (unsigned int)(v6->IdleTimeoutType - 1) <= 1 == ((p_m_TimeoutMgmt->m_IdleTimeoutStatus & 2) != 0) )
    goto LABEL_62;
  v25 = "should";
  if ( !(_DWORD)Settings )
    v25 = "should not";
  WPP_IFR_SF_sd(
    this->m_Globals,
    p_m_TimeoutMgmt->m_IdleTimeoutStatus & 2,
    0xCu,
    0x35u,
    (const _GUID *)&WPP_FxPkgPnp_cpp_Traceguids,
    v25,
    -1073741808);
  FxVerifierDbgBreakPoint(this->m_Globals);
  return -1073741808;
}