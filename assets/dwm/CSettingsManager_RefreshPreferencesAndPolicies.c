void __fastcall CSettingsManager::RefreshPreferencesAndPolicies(CSettingsManager *this)
{
  _DWORD *v1; // rdi
  int v3; // ebx
  __int64 v4; // r14
  __int64 v5; // r8
  int v6; // ecx
  int v7; // eax
  _DWORD *v8; // rdi
  int v9; // ebx
  __int64 v10; // r14
  __int64 v11; // r8
  int v12; // ecx
  int v13; // eax
  int v14; // [rsp+20h] [rbp-49h] BYREF
  const wchar_t *v15; // [rsp+28h] [rbp-41h]
  _DWORD v16[2]; // [rsp+30h] [rbp-39h] BYREF
  const wchar_t *v17; // [rsp+38h] [rbp-31h]
  int v18; // [rsp+40h] [rbp-29h]
  int v19; // [rsp+44h] [rbp-25h]
  const wchar_t *v20; // [rsp+48h] [rbp-21h]
  int v21; // [rsp+50h] [rbp-19h]
  int v22; // [rsp+54h] [rbp-15h]
  const wchar_t *v23; // [rsp+60h] [rbp-9h]
  _DWORD v24[2]; // [rsp+68h] [rbp-1h] BYREF
  const wchar_t *v25; // [rsp+70h] [rbp+7h]
  int v26; // [rsp+78h] [rbp+Fh]
  int v27; // [rsp+7Ch] [rbp+13h]
  const wchar_t *v28; // [rsp+80h] [rbp+17h]
  int v29; // [rsp+88h] [rbp+1Fh]
  int v30; // [rsp+8Ch] [rbp+23h]
  const wchar_t *v31; // [rsp+90h] [rbp+27h]
  int v32; // [rsp+98h] [rbp+2Fh]
  int v33; // [rsp+9Ch] [rbp+33h]

  v27 = 0;
  v30 = 0;
  v1 = v24;
  v33 = 0;
  v16[1] = 0;
  v19 = 0;
  v22 = 0;
  v3 = *((_DWORD *)this + 16);
  v23 = L"UseDPIScaling";
  v4 = 4LL;
  v24[0] = 1;
  v25 = L"AnimationsShiftKey";
  v28 = L"DisableLockingMemory";
  v31 = L"ModeChangeCurtainUseDebugColor";
  v15 = L"DisallowAnimations";
  v17 = L"DisallowColorizationColorChanges";
  v20 = L"DefaultColorizationColorState";
  v24[1] = 1;
  v26 = 2;
  v29 = 64;
  v32 = 128;
  v16[0] = 1;
  v18 = 2;
  v21 = 4;
  do
  {
    v5 = *((_QWORD *)v1 - 1);
    v14 = 0;
    if ( (int)CSettingsManager::GetDword(this, 0LL, v5, &v14) >= 0 )
      v6 = v14;
    else
      v6 = v1[1];
    v7 = *v1;
    if ( v6 )
      v3 |= v7;
    else
      v3 &= ~v7;
    v1 += 4;
    --v4;
  }
  while ( v4 );
  *((_DWORD *)this + 16) = v3;
  v8 = v16;
  v9 = *((_DWORD *)this + 17);
  v10 = 3LL;
  do
  {
    v11 = *((_QWORD *)v8 - 1);
    v14 = 0;
    if ( (int)CSettingsManager::GetDword(this, 1LL, v11, &v14) >= 0 )
      v12 = v14;
    else
      v12 = v8[1];
    v13 = *v8;
    if ( v12 )
      v9 |= v13;
    else
      v9 &= ~v13;
    v8 += 4;
    --v10;
  }
  while ( v10 );
  *((_DWORD *)this + 17) = v9;
}
