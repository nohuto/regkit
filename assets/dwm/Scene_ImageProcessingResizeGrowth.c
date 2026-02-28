__int64 dynamic_initializer_for__CCommonRegistryData::Scene::ImageProcessingResizeGrowth__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSceneSwitches",
             L"Software\\Microsoft\\Windows\\Dwm\\Scene",
             L"ImageProcessingResizeGrowth",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 200;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::Scene::ImageProcessingResizeGrowth = v1;
  return result;
}
