NTSTATUS __fastcall ReadManifestAssignedValue(_DWORD *a1)
{
  NTSTATUS result; // eax
  int v3; // edx
  int v4; // r9d
  HANDLE Handle; // [rsp+40h] [rbp+8h] BYREF

  Handle = 0LL;
  *a1 = 0;
  result = MyRegOpenKeyForRead(a1, L"\\Registry\\Machine\\SYSTEM\\CurrentControlSet\\Control\\USB", &Handle);
  if ( result < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_10;
    v4 = 17;
    goto LABEL_4;
  }
  result = MyRegQueryUlong(Handle);
  if ( result >= 0 )
  {
    if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      LOBYTE(v3) = 4;
      result = WPP_RECORDER_SF_d(
                 WPP_GLOBAL_Control->DeviceExtension,
                 v3,
                 1,
                 19,
                 (__int64)&WPP_5169c4c8089132207a438b4341aed5b6_Traceguids,
                 *a1);
    }
  }
  else if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    v4 = 18;
LABEL_4:
    LOBYTE(v3) = 2;
    result = WPP_RECORDER_SF_d(
               WPP_GLOBAL_Control->DeviceExtension,
               v3,
               1,
               v4,
               (__int64)&WPP_5169c4c8089132207a438b4341aed5b6_Traceguids,
               result);
  }
LABEL_10:
  if ( Handle )
    return ZwClose(Handle);
  return result;
}