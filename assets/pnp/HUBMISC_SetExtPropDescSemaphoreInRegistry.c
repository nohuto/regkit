__int64 __fastcall HUBMISC_SetExtPropDescSemaphoreInRegistry(__int64 a1)
{
  int v2; // edx
  int v3; // ebx
  int v4; // r9d
  int v5; // ecx
  int v7; // [rsp+48h] [rbp+10h] BYREF

  v7 = 1;
  v3 = HUBREG_WriteValueToDeviceHardwareKey(a1, (unsigned int)L"(*", 4, 4, (__int64)&v7); // ExtPropDescSemaphore
  if ( v3 >= 0 )
  {
    v7 = *(unsigned __int16 *)(a1 + 2000);
    v3 = HUBREG_WriteValueToDeviceHardwareKey(a1, (unsigned int)&g_RevisionId, 4, 4, (__int64)&v7);
    if ( v3 >= 0 )
    {
      if ( (*(_DWORD *)(a1 + 2464) & 0x400) != 0 )
        v5 = *(unsigned __int16 *)(*(_QWORD *)(a1 + 2528) + 4LL);
      else
        v5 = 0;
      v7 = v5;
      v3 = HUBREG_WriteValueToDeviceHardwareKey(a1, (unsigned int)&g_VendorRevision, 4, 4, (__int64)&v7);
      if ( v3 < 0 && WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        v4 = 63;
        goto LABEL_13;
      }
    }
    else if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      v4 = 62;
      goto LABEL_13;
    }
  }
  else if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    v4 = 61;
LABEL_13:
    LOBYTE(v2) = 2;
    WPP_RECORDER_SF_d(
      *(_QWORD *)(*(_QWORD *)(a1 + 8) + 1432LL),
      v2,
      5,
      v4,
      (__int64)&WPP_f96a94952a6932bc87af489d3d93d325_Traceguids,
      v3);
  }
  return ((v3 >> 31) & 0xFFFFFFF4) + 4077;
}