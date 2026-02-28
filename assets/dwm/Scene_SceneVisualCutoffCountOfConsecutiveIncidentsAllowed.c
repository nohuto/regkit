__int64 dynamic_initializer_for__CCommonRegistryData::Scene::SceneVisualCutoffCountOfConsecutiveIncidentsAllowed__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSceneSwitches",
             L"Software\\Microsoft\\Windows\\Dwm\\Scene",
             L"SceneVisualCutoffCountOfConsecutiveIncidentsAllowed",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 5;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::Scene::SceneVisualCutoffCountOfConsecutiveIncidentsAllowed = v1;
  return result;
}
