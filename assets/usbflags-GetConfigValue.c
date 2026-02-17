__int64 __fastcall GetConfigValue(__int64 a1, int a2, const void *a3, unsigned int a4, __int64 a5, _QWORD *a6)
{
  unsigned int v6; // ebx
  __int64 v9; // rcx
  void *Pool2; // rax
  void *v12; // rdi
  char v13; // [rsp+28h] [rbp-10h]

  v6 = 0;
  if ( a2 != 1 )
  {
    if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      v13 = a2;
      LOBYTE(a2) = 2;
      WPP_RECORDER_SF_d(g_RecorderLog, a2, 8, 10, (__int64)&WPP_c8c0b630977f39a4a91a92e2db87f7d2_Traceguids, v13);
    }
    return (unsigned int)-1073741811;
  }
  if ( !a4 )
    return (unsigned int)-1073741811;
  v9 = 256LL;
  if ( ExDefaultNonPagedPoolType == 1 )
    v9 = 64LL;
  Pool2 = (void *)ExAllocatePool2(v9, a4, 1130525525LL);
  v12 = Pool2;
  if ( Pool2 )
  {
    memset(Pool2, 0, a4);
    memmove(v12, a3, a4);
    *a6 = v12;
  }
  else
  {
    return (unsigned int)-1073741670;
  }
  return v6;
}