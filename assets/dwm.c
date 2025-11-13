__int64 dynamic_initializer_for__CCommonRegistryData::OverlayMinFPS__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"OverlayMinFPS",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 15;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::OverlayMinFPS = v1;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::InkGPUAccelOverrideVendorWhitelist__()
{
  bool v0; // bl
  __int64 result; // rax
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v0 = 0;
  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"InkGPUAccelOverrideVendorWhitelist",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  if ( !(_DWORD)result )
    v0 = v2 != 0;
  CCommonRegistryData::InkGPUAccelOverrideVendorWhitelist = v0;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::BlackOutAllReadback__()
{
  bool v0; // bl
  __int64 result; // rax
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v0 = 0;
  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"BlackOutAllReadback",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  if ( !(_DWORD)result )
    v0 = v2 != 0;
  CCommonRegistryData::BlackOutAllReadback = v0;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::DisplayChangeTimeoutMs__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"DisplayChangeTimeoutMs",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 1000;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::DisplayChangeTimeoutMs = v1;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::Scene::MsaaQualityMode__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSceneSwitches",
             L"Software\\Microsoft\\Windows\\Dwm\\Scene",
             L"MsaaQualityMode",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 2;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::Scene::MsaaQualityMode = v1;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::CpuClipAreaThreshold__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"CpuClipAreaThreshold",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 20000;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::CpuClipAreaThreshold = v1;
  return result;
}

bool dynamic_initializer_for__CCommonRegistryData::EnableEffectCaching__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"EnableEffectCaching",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::EnableEffectCaching = result;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::MegaRectSearchCount__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"MegaRectSearchCount",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 100;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::MegaRectSearchCount = v1;
  return result;
}

bool dynamic_initializer_for__CCommonRegistryData::OptimizeForDirtyExpressions__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"OptimizeForDirtyExpressions",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::OptimizeForDirtyExpressions = result;
  return result;
}

bool dynamic_initializer_for__CCommonRegistryData::Scene::EnableImageProcessing__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSceneSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm\\Scene",
                       L"EnableImageProcessing",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::Scene::EnableImageProcessing = result;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::ForceFullDirtyRendering__()
{
  bool v0; // bl
  __int64 result; // rax
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v0 = 0;
  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"ForceFullDirtyRendering",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  if ( !(_DWORD)result )
    v0 = v2 != 0;
  CCommonRegistryData::ForceFullDirtyRendering = v0;
  return result;
}

bool dynamic_initializer_for__CCommonRegistryData::GammaBlendWithFP16__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"GammaBlendWithFP16",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::GammaBlendWithFP16 = result;
  return result;
}

bool dynamic_initializer_for__CCommonRegistryData::EnableBackdropBlurCaching__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"EnableBackdropBlurCaching",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::EnableBackdropBlurCaching = result;
  return result;
}

bool dynamic_initializer_for__CCommonRegistryData::CpuClipAASinkEnableIntermediates__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"CpuClipAASinkEnableIntermediates",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::CpuClipAASinkEnableIntermediates = result;
  return result;
}

bool dynamic_initializer_for__CCommonRegistryData::EnableMegaRects__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"EnableMegaRects",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::EnableMegaRects = result;
  return result;
}

bool dynamic_initializer_for__CCommonRegistryData::EnableFrontBufferRenderChecks__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"EnableFrontBufferRenderChecks",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::EnableFrontBufferRenderChecks = result;
  return result;
}

bool dynamic_initializer_for__CCommonRegistryData::CpuClipAASinkEnableOcclusion__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"CpuClipAASinkEnableOcclusion",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::CpuClipAASinkEnableOcclusion = result;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::Scene::EnableBloom__()
{
  bool v0; // bl
  __int64 result; // rax
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v0 = 0;
  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSceneSwitches",
             L"Software\\Microsoft\\Windows\\Dwm\\Scene",
             L"EnableBloom",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  if ( !(_DWORD)result )
    v0 = v2 != 0;
  CCommonRegistryData::Scene::EnableBloom = v0;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::LogExpressionPerfStats__()
{
  bool v0; // bl
  __int64 result; // rax
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v0 = 0;
  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"LogExpressionPerfStats",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  if ( !(_DWORD)result )
    v0 = v2 != 0;
  CCommonRegistryData::LogExpressionPerfStats = v0;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::WarpEnableDebugColor__()
{
  bool v0; // bl
  __int64 result; // rax
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v0 = 0;
  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"WarpEnableDebugColor",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  if ( !(_DWORD)result )
    v0 = v2 != 0;
  CCommonRegistryData::WarpEnableDebugColor = v0;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::UseFastestMonitorAsPrimary__()
{
  bool v0; // bl
  __int64 result; // rax
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v0 = 0;
  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"UseFastestMonitorAsPrimary",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  if ( !(_DWORD)result )
    v0 = v2 != 0;
  CCommonRegistryData::UseFastestMonitorAsPrimary = v0;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::DisableProjectedShadows__()
{
  bool v0; // bl
  __int64 result; // rax
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v0 = 0;
  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"DisableProjectedShadows",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  if ( !(_DWORD)result )
    v0 = v2 != 0;
  CCommonRegistryData::DisableProjectedShadows = v0;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::DisableDrawListCaching__()
{
  bool v0; // bl
  __int64 result; // rax
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v0 = 0;
  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"DisableDrawListCaching",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  if ( !(_DWORD)result )
    v0 = v2 != 0;
  CCommonRegistryData::DisableDrawListCaching = v0;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::MegaRectSize__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"MegaRectSize",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 100000;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::MegaRectSize = v1;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::SuperWetExtensionTimeMicroseconds__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"SuperWetExtensionTimeMicroseconds",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 1000;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::SuperWetExtensionTimeMicroseconds = v1;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::RenderThreadTimeoutMilliseconds__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"RenderThreadTimeoutMilliseconds",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 5000;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::RenderThreadTimeoutMilliseconds = v1;
  return result;
}

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

bool dynamic_initializer_for__CCommonRegistryData::ConfigureInput__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"ConfigureInput",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::ConfigureInput = result;
  return result;
}

bool dynamic_initializer_for__CCommonRegistryData::ConfigureInput__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"ConfigureInput",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::ConfigureInput = result;
  return result;
}

bool dynamic_initializer_for__CCommonRegistryData::GammaBlendPencil__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"GammaBlendPencil",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::GammaBlendPencil = result;
  return result;
}

bool dynamic_initializer_for__CCommonRegistryData::Scene::EnableDrawToBackbuffer__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSceneSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm\\Scene",
                       L"EnableDrawToBackbuffer",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::Scene::EnableDrawToBackbuffer = result;
  return result;
}

bool dynamic_initializer_for__CCommonRegistryData::CpuClipAASinkEnableRender__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"CpuClipAASinkEnableRender",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::CpuClipAASinkEnableRender = result;
  return result;
}

bool dynamic_initializer_for__CCommonRegistryData::EnableCpuClipping__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"EnableCpuClipping",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::EnableCpuClipping = result;
  return result;
}

bool dynamic_initializer_for__CCommonRegistryData::EnablePrimitiveReordering__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"EnablePrimitiveReordering",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::EnablePrimitiveReordering = result;
  return result;
}

bool dynamic_initializer_for__CCommonRegistryData::UniformSpaceDpiMode__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"UniformSpaceDpiMode",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::UniformSpaceDpiMode = result;
  return result;
}

bool dynamic_initializer_for__CCommonRegistryData::EnableCommonSuperSets__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"EnableCommonSuperSets",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::EnableCommonSuperSets = result;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::Scene::SceneVisualCutoffThresholdInMS__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSceneSwitches",
             L"Software\\Microsoft\\Windows\\Dwm\\Scene",
             L"SceneVisualCutoffThresholdInMS",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 1000;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::Scene::SceneVisualCutoffThresholdInMS = v1;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::TelemetryFramesSequenceMaximumPeriodMilliseconds__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"TelemetryFramesSequenceMaximumPeriodMilliseconds",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 1000;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::TelemetryFramesSequenceMaximumPeriodMilliseconds = v1;
  return result;
}


__int64 dynamic_initializer_for__CCommonRegistryData::LayerClippingMode__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"LayerClippingMode",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 2;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::LayerClippingMode = v1;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::TelemetryFramesSequenceIdleIntervalMilliseconds__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"TelemetryFramesSequenceIdleIntervalMilliseconds",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 1000;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::TelemetryFramesSequenceIdleIntervalMilliseconds = v1;
  return result;
}

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

__int64 dynamic_initializer_for__CCommonRegistryData::SuperWetTiming::PeriodicFenceMinDifferenceMicroseconds__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"SuperWetTiming",
             L"Software\\Microsoft\\Windows\\Dwm\\GpuAccelInkTiming",
             L"PeriodicFenceMinDifferenceMicroseconds",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 500;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::SuperWetTiming::PeriodicFenceMinDifferenceMicroseconds = v1;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::SuperWetTiming::RefreshRatePercentage__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"SuperWetTiming",
             L"Software\\Microsoft\\Windows\\Dwm\\GpuAccelInkTiming",
             L"RefreshRatePercentage",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 10;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::SuperWetTiming::RefreshRatePercentage = v1;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::MajorityScreenTest_MinArea__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"MajorityScreenTest_MinArea",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 80;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::MajorityScreenTest_MinArea = v1;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::MousewheelAnimationDurationMs__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"MousewheelAnimationDurationMs",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 250;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::MousewheelAnimationDurationMs = v1;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::vBlankWaitTimeoutMonitorOffMs__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"vBlankWaitTimeoutMonitorOffMs",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 250;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::vBlankWaitTimeoutMonitorOffMs = v1;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::CpuClipWarpPartitionThreshold__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"CpuClipWarpPartitionThreshold",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 1024;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::CpuClipWarpPartitionThreshold = v1;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::SuperWetTiming::ExtensionTimeMicroseconds__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"SuperWetTiming",
             L"Software\\Microsoft\\Windows\\Dwm\\GpuAccelInkTiming",
             L"ExtensionTimeMicroseconds",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 1000;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::SuperWetTiming::ExtensionTimeMicroseconds = v1;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::MajorityScreenTest_MinLength__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"MajorityScreenTest_MinLength",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 80;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::MajorityScreenTest_MinLength = v1;
  return result;
}

__int64 dynamic_initializer_for__CCommonRegistryData::TelemetryFramesReportPeriodMilliseconds__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"TelemetryFramesReportPeriodMilliseconds",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 300000;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::TelemetryFramesReportPeriodMilliseconds = v1;
  return result;
}

void CCommonRegistryData::InitializeDWMKeysFromRegistry(void)
{
  unsigned int v0; // eax
  LONGLONG v1; // rcx
  LONGLONG v2; // rdx
  int v3; // eax
  unsigned int v4; // [rsp+70h] [rbp+30h] BYREF

  v4 = 0;
  if ( !(unsigned int)GetPersistedRegistryValueW(
                        L"DWMSwitches",
                        L"Software\\Microsoft\\Windows\\Dwm",
                        L"OverlayTestMode",
                        16LL,
                        0LL,
                        &v4,
                        4,
                        0LL) )
    CCommonRegistryData::m_dwOverlayTestMode = v4;
  v4 = 0;
  GetPersistedRegistryValueW(
    L"DWMSwitches",
    L"Software\\Microsoft\\Windows\\Dwm",
    L"DisableAdvancedDirectFlip",
    16LL,
    0LL,
    &v4,
    4,
    0LL);
  v4 = 0;
  if ( !(unsigned int)GetPersistedRegistryValueW(
                        L"DWMSwitches",
                        L"Software\\Microsoft\\Windows\\Dwm",
                        L"DisableIndependentFlip",
                        16LL,
                        0LL,
                        &v4,
                        4,
                        0LL) )
    CCommonRegistryData::m_fDisableIndependentFlip = v4 != 0;
  v4 = 0;
  if ( !(unsigned int)GetPersistedRegistryValueW(
                        L"DWMSwitches",
                        L"Software\\Microsoft\\Windows\\Dwm",
                        L"FrameCounterPosition",
                        16LL,
                        0LL,
                        &v4,
                        4,
                        0LL) )
    CCommonRegistryData::m_fDebugFrameCounterIsVertical = v4 != 0;
  v4 = 0;
  if ( !(unsigned int)GetPersistedRegistryValueW(
                        L"DWMSwitches",
                        L"Software\\Microsoft\\Windows\\Dwm",
                        L"FlattenVirtualSurfaceEffectInput",
                        16LL,
                        0LL,
                        &v4,
                        4,
                        0LL) )
    CCommonRegistryData::m_fFlattenVirtualSurfaceBrush = v4 != 0;
  if ( !(unsigned int)GetPersistedRegistryValueW(
                        L"DWMSwitches",
                        L"Software\\Microsoft\\Windows\\Dwm",
                        L"CpuClipFlatteningTolerance",
                        16LL,
                        0LL,
                        &v4,
                        4,
                        0LL) )
    CCommonRegistryData::m_flCpuClipFlatteningTolerance = (float)(int)v4 / 1000.0;
  if ( !(unsigned int)GetPersistedRegistryValueW(
                        L"DWMSwitches",
                        L"Software\\Microsoft\\Windows\\Dwm",
                        L"InteractionOutputPredictionDisabled",
                        16LL,
                        0LL,
                        &v4,
                        4,
                        0LL) )
    CCommonRegistryData::m_fDisableInteractionOutputPrediction = v4 != 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"BackdropBlurCachingThrottleMs",
                       16LL,
                       0LL,
                       &v4,
                       4,
                       0LL) )
  {
    v1 = 25 * g_qpcFrequency.QuadPart;
  }
  else
  {
    v0 = v4;
    if ( v4 > 0x3E8 )
      v0 = 1000;
    v1 = g_qpcFrequency.QuadPart * v0;
  }
  v4 = 0;
  CCommonRegistryData::m_backdropBlurCachingThrottleQPCTimeDelta = v1 / 1000;
  if ( !(unsigned int)GetPersistedRegistryValueW(
                        L"DWMSceneSwitches",
                        L"Software\\Microsoft\\Windows\\Dwm\\Scene",
                        L"ForceNonPrimaryDisplayAdapter",
                        16LL,
                        0LL,
                        &v4,
                        4,
                        0LL) )
    CCommonRegistryData::m_fSceneForceNonPrimaryDisplayAdapter = v4 != 0;
  if ( !(unsigned int)GetPersistedRegistryValueW(
                        L"DWMSceneSwitches",
                        L"Software\\Microsoft\\Windows\\Dwm\\Scene",
                        L"ImageProcessingResizeThreshold",
                        16LL,
                        0LL,
                        &v4,
                        4,
                        0LL) )
    CCommonRegistryData::m_flSceneImageProcessingResizeThreshold = (float)(int)v4 / 100.0;
  v4 = 0;
  if ( !(unsigned int)GetPersistedRegistryValueW(
                        L"DWMSwitches",
                        L"Software\\Microsoft\\Windows\\Dwm",
                        L"ForceEffectMode",
                        16LL,
                        0LL,
                        &v4,
                        4,
                        0LL)
    && v4 <= 2 )
  {
    CCommonRegistryData::m_forceEffectMode = v4;
  }
  CCommonRegistryData::m_compositorClockPolicy = 1;
  v4 = 1;
  if ( !(unsigned int)GetPersistedRegistryValueW(
                        L"DWMSwitches",
                        L"Software\\Microsoft\\Windows\\Dwm",
                        L"CompositorClockPolicy",
                        16LL,
                        0LL,
                        &v4,
                        4,
                        0LL)
    && v4 < 2 )
  {
    CCommonRegistryData::m_compositorClockPolicy = v4;
  }
  v4 = 1;
  if ( !(unsigned int)GetPersistedRegistryValueW(
                        L"DWMSwitches",
                        L"Software\\Microsoft\\Windows\\Dwm",
                        L"ParallelModePolicy",
                        16LL,
                        0LL,
                        &v4,
                        4,
                        0LL) )
  {
    v3 = v4;
    if ( v4 >= 3 )
      v3 = 1;
    CCommonRegistryData::m_parallelModePolicy = v3;
  }
  v4 = 1;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"ParallelModeRateThreshold",
                       16LL,
                       0LL,
                       &v4,
                       4,
                       0LL) )
  {
    v2 = g_qpcFrequency.QuadPart / 119;
  }
  else if ( v4 )
  {
    v2 = g_qpcFrequency.QuadPart / v4;
  }
  else
  {
    v2 = 0LL;
  }
  CCommonRegistryData::m_parallelModeDurationThreshold = v2;
  v4 = 0;
  if ( !(unsigned int)GetPersistedRegistryValueW(
                        L"DWMSwitches",
                        L"Software\\Microsoft\\Windows\\Dwm",
                        L"CustomRefreshRateMode",
                        16LL,
                        0LL,
                        &v4,
                        4,
                        0LL)
    && v4 <= 2 )
  {
    CCommonRegistryData::m_customRefreshRateMode = v4;
  }
  v4 = 0;
  if ( !(unsigned int)GetPersistedRegistryValueW(
                        L"DWMSwitches",
                        L"Software\\Microsoft\\Windows\\Dwm",
                        L"SDRBoostPercentOverride",
                        16LL,
                        0LL,
                        &v4,
                        4,
                        0LL) )
    CCommonRegistryData::m_flSDRBoostOverride = (float)(int)v4 / 100.0;
  v4 = 0;
  GetPersistedRegistryValueW(
    L"DWMSwitches",
    L"Software\\Microsoft\\Windows\\Dwm",
    L"DisableProjectedShadowsRendering",
    16LL,
    0LL,
    &v4,
    4,
    0LL);
  v4 = 0;
  if ( !(unsigned int)GetPersistedRegistryValueW(
                        L"DWMSwitches",
                        L"Software\\Microsoft\\Windows\\Dwm",
                        L"ShowDirtyRegions",
                        16LL,
                        0LL,
                        &v4,
                        4,
                        0LL) )
    CCommonRegistryData::m_fShowDirtyRegions = v4;
  v4 = 0;
  if ( !(unsigned int)GetPersistedRegistryValueW(
                        L"DWMSwitches",
                        L"Software\\Microsoft\\Windows\\Dwm",
                        L"ResampleModeOverride",
                        16LL,
                        0LL,
                        &v4,
                        4,
                        0LL) )
    CCommonRegistryData::m_dwResampleModeOverride = v4;
  v4 = 0;
  if ( !(unsigned int)GetPersistedRegistryValueW(
                        L"DWMSwitches",
                        L"Software\\Microsoft\\Windows\\Dwm",
                        L"ResampleInLinearSpace",
                        16LL,
                        0LL,
                        &v4,
                        4,
                        0LL) )
    CCommonRegistryData::m_fResampleInLinearSpace = v4;
}

// win32kfull.sys

_BOOL8 bDwmChildWindowDpiIsolationEnabled(void)
{
  BOOL v0; // ebx
  ULONG ResultLength; // [rsp+30h] [rbp-D0h] BYREF
  void *KeyHandle; // [rsp+38h] [rbp-C8h] BYREF
  struct _UNICODE_STRING DestinationString; // [rsp+40h] [rbp-C0h] BYREF
  struct _OBJECT_ATTRIBUTES ObjectAttributes; // [rsp+50h] [rbp-B0h] BYREF
  _BYTE KeyValueInformation[4]; // [rsp+80h] [rbp-80h] BYREF
  int v7; // [rsp+84h] [rbp-7Ch]
  int v8; // [rsp+8Ch] [rbp-74h]

  *(&ObjectAttributes.Length + 1) = 0;
  *(&ObjectAttributes.Attributes + 1) = 0;
  KeyHandle = 0LL;
  ResultLength = 0;
  v0 = 1;
  DestinationString = 0LL;
  RtlInitUnicodeString(&DestinationString, L"\\Registry\\Machine\\Software\\Microsoft\\Windows\\DWM");
  ObjectAttributes.RootDirectory = 0LL;
  ObjectAttributes.ObjectName = &DestinationString;
  ObjectAttributes.Length = 48;
  ObjectAttributes.Attributes = 576;
  *(_OWORD *)&ObjectAttributes.SecurityDescriptor = 0LL;
  if ( ZwOpenKey(&KeyHandle, 0x20019u, &ObjectAttributes) >= 0 )
  {
    RtlInitUnicodeString(&DestinationString, L"ChildWindowDpiIsolation");
    if ( ZwQueryValueKey(
           KeyHandle,
           &DestinationString,
           KeyValuePartialInformation,
           KeyValueInformation,
           0x400u,
           &ResultLength) >= 0
      && v7 == 4 )
    {
      v0 = v8 != 0;
    }
    ZwClose(KeyHandle);
  }
  return v0;
}

__int64 __fastcall bDwmResizeOptimizationOverride(unsigned int *a1, unsigned int *a2, unsigned int *a3)
{
  unsigned int v3; // ebx
  ULONG ResultLength; // [rsp+30h] [rbp-D0h] BYREF
  void *KeyHandle; // [rsp+38h] [rbp-C8h] BYREF
  struct _UNICODE_STRING DestinationString; // [rsp+40h] [rbp-C0h] BYREF
  struct _OBJECT_ATTRIBUTES ObjectAttributes; // [rsp+50h] [rbp-B0h] BYREF
  _BYTE KeyValueInformation[4]; // [rsp+80h] [rbp-80h] BYREF
  int v13; // [rsp+84h] [rbp-7Ch]
  unsigned int v14; // [rsp+8Ch] [rbp-74h]

  v3 = 0;
  *(&ObjectAttributes.Length + 1) = 0;
  *(&ObjectAttributes.Attributes + 1) = 0;
  KeyHandle = 0LL;
  ResultLength = 0;
  DestinationString = 0LL;
  RtlInitUnicodeString(&DestinationString, L"\\Registry\\Machine\\Software\\Microsoft\\Windows\\DWM");
  ObjectAttributes.RootDirectory = 0LL;
  ObjectAttributes.ObjectName = &DestinationString;
  ObjectAttributes.Length = 48;
  ObjectAttributes.Attributes = 576;
  *(_OWORD *)&ObjectAttributes.SecurityDescriptor = 0LL;
  if ( ZwOpenKey(&KeyHandle, 0x20019u, &ObjectAttributes) >= 0 )
  {
    RtlInitUnicodeString(&DestinationString, L"EnableResizeOptimization");
    if ( ZwQueryValueKey(
           KeyHandle,
           &DestinationString,
           KeyValuePartialInformation,
           KeyValueInformation,
           0x400u,
           &ResultLength) >= 0
      && v13 == 4 )
    {
      v3 = 1;
      *a1 = v14;
    }
    RtlInitUnicodeString(&DestinationString, L"ResizeTimeoutGdi");
    if ( ZwQueryValueKey(
           KeyHandle,
           &DestinationString,
           KeyValuePartialInformation,
           KeyValueInformation,
           0x400u,
           &ResultLength) >= 0
      && v13 == 4 )
    {
      *a2 = v14;
    }
    RtlInitUnicodeString(&DestinationString, L"ResizeTimeoutModern");
    if ( ZwQueryValueKey(
           KeyHandle,
           &DestinationString,
           KeyValuePartialInformation,
           KeyValueInformation,
           0x400u,
           &ResultLength) >= 0
      && v13 == 4 )
    {
      *a3 = v14;
    }
    ZwClose(KeyHandle);
  }
  return v3;
}

_BOOL8 bDwmDeviceBitmapsEnabled(void)
{
  BOOL v0; // ebx
  ULONG ResultLength; // [rsp+30h] [rbp-D0h] BYREF
  void *KeyHandle; // [rsp+38h] [rbp-C8h] BYREF
  struct _UNICODE_STRING DestinationString; // [rsp+40h] [rbp-C0h] BYREF
  struct _OBJECT_ATTRIBUTES ObjectAttributes; // [rsp+50h] [rbp-B0h] BYREF
  _BYTE KeyValueInformation[4]; // [rsp+80h] [rbp-80h] BYREF
  int v7; // [rsp+84h] [rbp-7Ch]
  int v8; // [rsp+8Ch] [rbp-74h]

  *(&ObjectAttributes.Length + 1) = 0;
  *(&ObjectAttributes.Attributes + 1) = 0;
  DestinationString = 0LL;
  KeyHandle = 0LL;
  v0 = 1;
  ResultLength = 0;
  RtlInitUnicodeString(&DestinationString, L"\\Registry\\Machine\\Software\\Microsoft\\Windows\\DWM");
  ObjectAttributes.Length = 48;
  ObjectAttributes.ObjectName = &DestinationString;
  ObjectAttributes.RootDirectory = 0LL;
  ObjectAttributes.Attributes = 576;
  *(_OWORD *)&ObjectAttributes.SecurityDescriptor = 0LL;
  if ( ZwOpenKey(&KeyHandle, 0x20019u, &ObjectAttributes) >= 0 )
  {
    RtlInitUnicodeString(&DestinationString, L"DisableDeviceBitmaps");
    if ( ZwQueryValueKey(
           KeyHandle,
           &DestinationString,
           KeyValuePartialInformation,
           KeyValueInformation,
           0x400u,
           &ResultLength) >= 0
      && v7 == 4 )
    {
      v0 = v8 == 0;
    }
    ZwClose(KeyHandle);
  }
  return v0;
}

// dwm.exe

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