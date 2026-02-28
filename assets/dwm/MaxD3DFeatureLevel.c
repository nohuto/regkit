__int64 dynamic_initializer_for__CCommonRegistryData::MaxD3DFeatureLevel__()
{
  int v0; // ebx
  __int64 result; // rax
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v0 = 0;
  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"MaxD3DFeatureLevel",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  if ( !(_DWORD)result )
    v0 = v2;
  CCommonRegistryData::MaxD3DFeatureLevel = v0;
  return result;
}
