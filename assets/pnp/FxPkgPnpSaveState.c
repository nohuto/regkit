void __fastcall FxPkgPnp::SaveState(FxPkgPnp *this, unsigned __int8 UseCanSaveState)
{
  FxPowerPolicyOwnerSettings *m_Owner; // rax
  const void *_a1; // rax
  _FX_DRIVER_GLOBALS *v5; // r10
  _IRP *m_PendingDevicePowerIrp; // rax
  FxPowerPolicyOwnerSettings *v7; // rax
  FxPowerPolicyOwnerSettings *v8; // rax
  _UNICODE_STRING name; // [rsp+30h] [rbp-10h] BYREF
  int Data; // [rsp+50h] [rbp+10h] BYREF
  FxAutoRegKey hKey; // [rsp+60h] [rbp+20h] BYREF

  m_Owner = this->m_PowerPolicyMachine.m_Owner;
  hKey.m_Key = 0LL;
  name = 0LL;
  if ( m_Owner )
  {
    if ( !UseCanSaveState || m_Owner->m_CanSaveState )
    {
      if ( (m_Owner->m_IdleSettings.Dirty || m_Owner->m_WakeSettings.Dirty)
        && (m_Owner->m_IdleSettings.Overridable || m_Owner->m_WakeSettings.Overridable)
        && (!this->m_SpecialSupport[0]
         || (m_PendingDevicePowerIrp = this->m_PendingDevicePowerIrp) == 0LL
         || m_PendingDevicePowerIrp->Tail.Overlay.CurrentStackLocation->Parameters.Read.ByteOffset.LowPart != 1)
        && FxDevice::OpenSettingsKey(this->m_Device, &hKey.m_Key, 0x20000u) >= 0 )
      {
        v7 = this->m_PowerPolicyMachine.m_Owner;
        if ( v7->m_IdleSettings.Overridable && v7->m_IdleSettings.Dirty )
        {
          RtlInitUnicodeString(&name, L"IdleInWorkingState");
          Data = this->m_PowerPolicyMachine.m_Owner->m_IdleSettings.Enabled;
          ZwSetValueKey(hKey.m_Key, &name, 0, 4u, &Data, 4u);
          this->m_PowerPolicyMachine.m_Owner->m_IdleSettings.Dirty = 0;
        }
        v8 = this->m_PowerPolicyMachine.m_Owner;
        if ( v8->m_WakeSettings.Overridable && v8->m_WakeSettings.Dirty )
        {
          RtlInitUnicodeString(&name, L"WakeFromSleepState");
          Data = this->m_PowerPolicyMachine.m_Owner->m_WakeSettings.Enabled;
          ZwSetValueKey(hKey.m_Key, &name, 0, 4u, &Data, 4u);
          this->m_PowerPolicyMachine.m_Owner->m_WakeSettings.Dirty = 0;
        }
      }
    }
    else
    {
      if ( !this->m_Globals->FxVerboseOn )
        return;
      _a1 = FxObject::GetObjectHandleUnchecked(this->m_DeviceBase);
      WPP_IFR_SF_q(v5, 5u, 0xCu, 0x50u, (const _GUID *)&WPP_FxPkgPnp_cpp_Traceguids, _a1);
    }
    if ( hKey.m_Key )
      ZwClose(hKey.m_Key);
  }
}