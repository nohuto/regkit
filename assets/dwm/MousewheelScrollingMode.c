__int64 dynamic_initializer_for__CCommonRegistryData::MousewheelScrollingMode__()
{
  int v0; // ebx
  __int64 result; // rax
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v0 = 0;
  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"MousewheelScrollingMode",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  if ( !(_DWORD)result )
    v0 = v2;
  CCommonRegistryData::MousewheelScrollingMode = v0;
  return result;
}
