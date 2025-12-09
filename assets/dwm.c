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

bool dynamic_initializer_for__CCommonRegistryData::EnableDDisplayScanoutCaching__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"EnableDDisplayScanoutCaching",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::EnableDDisplayScanoutCaching = result;
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

// dwminit.dll

__int64 __fastcall DwmpCreateSessionProcess(PVOID Parameter)
{
  unsigned __int64 v1; // r14
  signed int v2; // ebx
  HANDLE Thread; // rdi
  int v4; // ecx
  int v5; // r8d
  int v6; // r9d
  signed int LastError; // eax
  GUID *v9; // [rsp+50h] [rbp-10h] BYREF
  __int64 v10; // [rsp+58h] [rbp-8h] BYREF
  int pvData; // [rsp+98h] [rbp+38h] BYREF
  DWORD pcbData; // [rsp+A0h] [rbp+40h] BYREF
  GUID *v13; // [rsp+A8h] [rbp+48h] BYREF

  v1 = (unsigned int)Parameter;
  v2 = 0;
  Thread = 0LL;
  if ( GetModuleHandleW(L"wininit.exe")
    && (pvData = 0,
        pcbData = 4,
        RegGetValueW(
          HKEY_LOCAL_MACHINE,
          L"Software\\Microsoft\\Windows\\DWM",
          L"OneCoreNoBootDWM",
          0x20000010u,
          0LL,
          &pvData,
          &pcbData),
        pvData) )
  {
    v2 = 1;
  }
  else if ( gDwmFirstLaunch )
  {
    SetLastError(0);
    Thread = CreateThread(0LL, 0LL, DwmpCreateSessionProcessWorker, (LPVOID)v1, 0, 0LL);
    if ( !Thread )
    {
      LastError = GetLastError();
      v2 = LastError;
      if ( LastError > 0 )
        v2 = (unsigned __int16)LastError | 0x80070000;
      if ( v2 >= 0 )
        v2 = -2003304445;
      DoStackCaptureDirect(v2, 0x5C7u);
    }
  }
  else
  {
    DwmpCreateSessionProcessWorker((PVOID)v1);
  }
  if ( (unsigned int)dword_180016000 > 5
    && (qword_180016010 & 0x400000000000LL) != 0
    && (qword_180016018 & 0x400000000000LL) == qword_180016018 )
  {
    pvData = v1;
    v13 = &gDwmInitTargetAppSessionGuid;
    pcbData = v2;
    v9 = &gDwmInitTelemetryActivityId;
    v10 = 0x1000000LL;
    _tlgWriteTemplate<long (_tlgProvider_t const *,void const *,_GUID const *,_GUID const *,unsigned int,_EVENT_DATA_DESCRIPTOR *),&long _tlgWriteTransfer_EventWriteTransfer(_tlgProvider_t const *,void const *,_GUID const *,_GUID const *,unsigned int,_EVENT_DATA_DESCRIPTOR *),_GUID const *,_GUID const *>::Write<_tlgWrapperByVal<8>,_tlgWrapperByRef<16>,_tlgWrapperByVal<4>,_tlgWrapperByVal<4>,_tlgWrapperByRef<16>>(
      v4,
      (unsigned int)&unk_180012317,
      v5,
      v6,
      (__int64)&v10,
      (__int64)&v9,
      (__int64)&pcbData,
      (__int64)&pvData,
      (__int64)&v13);
  }
  if ( Thread )
    CloseHandle(Thread);
  return (unsigned int)v2;
}

// uDWM.dll

// Hidden C++ exception states: #wind=3
__int64 __fastcall CDesktopManager::Initialize(CDesktopManager *this, struct IUnknown *a2)
{
  HANDLE EventW; // r15
  int v4; // eax
  signed int v5; // ebx
  __int64 v6; // rax
  bool v7; // r14
  __int64 v8; // rax
  LSTATUS v9; // ebx
  CWindowList *v10; // rax
  CWindowList *v11; // rbx
  CWindowList *v12; // rax
  int v13; // r9d
  CContactManager *v14; // rax
  CContactManager *v15; // rax
  CTransitionVisualController *v16; // rax
  const struct std::nothrow_t *v17; // rdx
  CTransitionVisualController *v18; // r11
  CAnimationScheduler *v19; // rax
  const struct std::nothrow_t *v20; // rdx
  CAnimationScheduler *v21; // r10
  CAnimationClockCoordinator *v22; // rax
  CAnimationClockCoordinator *v23; // r8
  CDesktopManager *v24; // rax
  int Factory; // eax
  CBaseObject *v26; // rcx
  CIconicBitmapRegistry *v27; // rax
  CIconicBitmapRegistry *v28; // rax
  CImmersiveIconicBitmapRegistry *v29; // rax
  CImmersiveIconicBitmapRegistry *v30; // rax
  __int64 v31; // rcx
  const struct std::nothrow_t *v32; // rdx
  CProjectionBorderManager *v33; // rax
  CProjectionBorderManager *v34; // rax
  __int64 v35; // rax
  signed int LastError; // eax
  int v37; // r9d
  HANDLE Thread; // rax
  signed int v39; // eax
  signed int v40; // eax
  unsigned int phkResult; // [rsp+20h] [rbp-50h]
  unsigned int phkResulta; // [rsp+20h] [rbp-50h]
  DWORD Type; // [rsp+30h] [rbp-40h] BYREF
  HKEY hKey; // [rsp+38h] [rbp-38h] BYREF
  __int64 v46; // [rsp+40h] [rbp-30h] BYREF
  __int64 pvParam; // [rsp+48h] [rbp-28h] BYREF
  HANDLE Handles[4]; // [rsp+50h] [rbp-20h] BYREF
  int v49; // [rsp+B8h] [rbp+48h] BYREF
  struct _RTL_CRITICAL_SECTION *Data; // [rsp+C0h] [rbp+50h] BYREF
  DWORD cbData; // [rsp+C8h] [rbp+58h] BYREF

  v46 = 0LL;
  EventW = 0LL;
  hKey = 0LL;
  v4 = ((__int64 (__fastcall *)(struct IUnknown *, GUID *, __int64 *))a2->lpVtbl->QueryInterface)(
         a2,
         &GUID_3ae5dff1_7681_484a_956a_6fd06c8e671e,
         &v46);
  v5 = v4;
  if ( v4 < 0 )
  {
    phkResult = 315;
    goto LABEL_108;
  }
  *((_BYTE *)this + 20) = 0;
  *((_BYTE *)this + 25) = 0;
  v6 = (*(__int64 (__fastcall **)(__int64))(*(_QWORD *)v46 + 24LL))(v46);
  *((_QWORD *)this + 9) = v6;
  v49 = 0;
  if ( (*(int (__fastcall **)(__int64, const wchar_t *, int *))(*(_QWORD *)v6 + 8LL))(
         v6,
         L"DefaultRemoteAppCreation",
         &v49) >= 0 )
    CDesktopManager::s_defaultRemoteAppCreation = v49 != 0;
  v7 = 0;
  v8 = wil::out_param<wil::unique_any_t<wil::details::unique_storage<wil::details::resource_policy<HKEY__ *,long (*)(HKEY__ *),&long RegCloseKey(HKEY__ *),wistd::integral_constant<unsigned __int64,0>,HKEY__ *,HKEY__ *,0,std::nullptr_t>>>>(
         Handles,
         &hKey);
  v9 = RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"Software\\Microsoft\\Windows\\Dwm", 0, 0x20019u, (PHKEY)(v8 + 8));
  wil::details::out_param_t<wil::unique_any_t<wil::details::unique_storage<wil::details::resource_policy<HKEY__ *,long (*)(HKEY__ *),&long RegCloseKey(HKEY__ *),wistd::integral_constant<unsigned __int64,0>,HKEY__ *,HKEY__ *,0,std::nullptr_t>>>>::~out_param_t<wil::unique_any_t<wil::details::unique_storage<wil::details::resource_policy<HKEY__ *,long (*)(HKEY__ *),&long RegCloseKey(HKEY__ *),wistd::integral_constant<unsigned __int64,0>,HKEY__ *,HKEY__ *,0,std::nullptr_t>>>>(Handles);
  if ( !v9 )
  {
    LODWORD(Data) = 0;
    cbData = 4;
    Type = 0;
    if ( !RegQueryValueExW(hKey, L"ForceEffectMode", 0LL, &Type, (LPBYTE)&Data, &cbData) && (_DWORD)Data == 2 )
      *((_BYTE *)this + 29) = 1;
    LODWORD(Data) = 0;
    cbData = 4;
    Type = 0;
    if ( !RegQueryValueExW(hKey, L"ForceUDwmSoftwareDevice", 0LL, &Type, (LPBYTE)&Data, &cbData) && (_DWORD)Data )
      v7 = 1;
    LODWORD(Data) = 0;
    cbData = 4;
    Type = 0;
    if ( !RegQueryValueExW(hKey, L"ForceDisableModeChangeAnimation", 0LL, &Type, (LPBYTE)&Data, &cbData) && (_DWORD)Data )
      CDesktopManager::s_forceDisableModeChangeAnimation = 1;
  }
  v10 = (CWindowList *)DefaultHeap::AllocClear(0x2C0uLL);
  v11 = v10;
  Data = (struct _RTL_CRITICAL_SECTION *)v10;
  if ( v10 )
  {
    memset_0(v10, 0, 0x2C0uLL);
    v12 = CWindowList::CWindowList(v11);
  }
  else
  {
    v12 = 0LL;
  }
  *((_QWORD *)this + 53) = v12;
  if ( v12 )
  {
    v14 = (CContactManager *)DefaultHeap::AllocClear(0x148uLL);
    Data = (struct _RTL_CRITICAL_SECTION *)v14;
    if ( v14 )
      v15 = CContactManager::CContactManager(v14);
    else
      v15 = 0LL;
    *((_QWORD *)this + 20) = v15;
    if ( !v15 )
    {
      phkResult = 396;
      goto LABEL_19;
    }
    v16 = (CTransitionVisualController *)DefaultHeap::AllocClear(0xC0uLL);
    Data = (struct _RTL_CRITICAL_SECTION *)v16;
    if ( v16 )
      v18 = CTransitionVisualController::CTransitionVisualController(v16);
    else
      v18 = 0LL;
    *((_QWORD *)CDesktopManager::s_pDesktopManagerInstance + 24) = v18;
    if ( !v18 )
    {
      phkResult = 399;
      goto LABEL_19;
    }
    v19 = (CAnimationScheduler *)operator new[](0x58uLL, v17);
    Data = (struct _RTL_CRITICAL_SECTION *)v19;
    if ( v19 )
      v21 = CAnimationScheduler::CAnimationScheduler(v19);
    else
      v21 = 0LL;
    *((_QWORD *)CDesktopManager::s_pDesktopManagerInstance + 23) = v21;
    if ( !v21 )
    {
      phkResult = 402;
      goto LABEL_19;
    }
    v22 = (CAnimationClockCoordinator *)operator new[](0x58uLL, v20);
    Data = (struct _RTL_CRITICAL_SECTION *)v22;
    if ( v22 )
      v23 = CAnimationClockCoordinator::CAnimationClockCoordinator(v22);
    else
      v23 = 0LL;
    v24 = CDesktopManager::s_pDesktopManagerInstance;
    *((_QWORD *)CDesktopManager::s_pDesktopManagerInstance + 21) = v23;
    if ( !v23 )
    {
      phkResult = 405;
      goto LABEL_19;
    }
    v4 = CAnimationClockCoordinator::SetEventCallback(
           v23,
           (struct IAnimationClockEventListener *)((*((_QWORD *)v24 + 23) + 8LL) & ((unsigned __int128)-(__int128)*((unsigned __int64 *)v24 + 23) >> 64)));
    v5 = v4;
    if ( v4 < 0 )
    {
      phkResult = 408;
    }
    else
    {
      pvParam = 8LL;
      if ( SystemParametersInfoW(0x48u, 8u, &pvParam, 0) )
        CDesktopManager::SetWindowAnimation(HIDWORD(pvParam) != 0);
      v4 = DwmRedirectionManagerInitialize(
             *((struct IDwmRedirectionClient **)this + 53),
             (struct IDwmRedirectionManager **)this + 8);
      v5 = v4;
      if ( v4 >= 0 )
      {
        wil::com_ptr_t<IDCompositionAnimationStats,wil::err_returncode_policy>::reset((char *)this + 48);
        Factory = CCompositor::Create((struct CCompositor **)this + 6);
        v5 = Factory;
        if ( Factory < 0 )
        {
          phkResulta = 426;
        }
        else
        {
          v26 = (CBaseObject *)*((_QWORD *)this + 7);
          *((_QWORD *)this + 7) = 0LL;
          if ( v26 )
            CBaseObject::Release(v26);
          Factory = CGraphicsDeviceManager::Create(v7, (struct CGraphicsDeviceManager **)this + 7);
          v5 = Factory;
          if ( Factory < 0 )
          {
            phkResulta = 429;
          }
          else
          {
            v27 = (CIconicBitmapRegistry *)DefaultHeap::AllocClear(0x70uLL);
            Data = (struct _RTL_CRITICAL_SECTION *)v27;
            if ( v27 )
              v28 = CIconicBitmapRegistry::CIconicBitmapRegistry(v27);
            else
              v28 = 0LL;
            *((_QWORD *)this + 28) = v28;
            if ( !v28 )
            {
              v5 = -2147024882;
              MilInstrumentationCheckHR_MaybeFailFast(0x14u, &dword_1800FD698, 1u, -2147024882, 0x1B5u, 0LL);
              goto LABEL_103;
            }
            Factory = CIconicBitmapRegistry::Init(v28);
            v5 = Factory;
            if ( Factory < 0 )
            {
              phkResulta = 438;
            }
            else
            {
              v29 = (CImmersiveIconicBitmapRegistry *)DefaultHeap::AllocClear(0x58uLL);
              Data = (struct _RTL_CRITICAL_SECTION *)v29;
              if ( v29 )
                v30 = CImmersiveIconicBitmapRegistry::CImmersiveIconicBitmapRegistry(v29);
              else
                v30 = 0LL;
              *((_QWORD *)this + 29) = v30;
              if ( !v30 )
              {
                v5 = -2147024882;
                MilInstrumentationCheckHR_MaybeFailFast(0x14u, &dword_1800FD698, 1u, -2147024882, 0x1B9u, 0LL);
                goto LABEL_103;
              }
              Factory = CImmersiveIconicBitmapRegistry::Init(v30);
              v5 = Factory;
              if ( Factory < 0 )
              {
                phkResulta = 442;
              }
              else
              {
                CDesktopManager::SetupDPIValues(this);
                *((_DWORD *)this + 118) = -1;
                if ( (Microsoft_Windows_Dwm_UdwmEnableBits & 1) != 0 )
                  McTemplateU0q_EtwEventWriteTransfer(v31, &UdwmStartup_Info, 1LL);
                CDesktopManager::UpdateRemotingMode(this);
                Factory = WICCreateImagingFactory_Proxy(567LL, (char *)this + 240);
                v5 = Factory;
                if ( Factory < 0 )
                {
                  phkResulta = 453;
                }
                else
                {
                  Factory = DWriteCreateFactory(0LL, &GUID_b859ee5a_d838_4b5b_a2e8_1adc7d93db48, (char *)this + 248);
                  v5 = Factory;
                  if ( Factory < 0 )
                  {
                    phkResulta = 460;
                  }
                  else
                  {
                    LODWORD(Data) = 13;
                    Factory = CDesktopManager::UpdateSettings(this, (unsigned int *)&Data);
                    v5 = Factory;
                    if ( Factory < 0 )
                    {
                      phkResulta = 475;
                    }
                    else
                    {
                      Factory = CLivePreview::Create((struct CLivePreview **)this + 57);
                      v5 = Factory;
                      if ( Factory < 0 )
                      {
                        phkResulta = 477;
                      }
                      else
                      {
                        v33 = (CProjectionBorderManager *)operator new[](0x290uLL, v32);
                        Data = (struct _RTL_CRITICAL_SECTION *)v33;
                        if ( v33 )
                          v34 = CProjectionBorderManager::CProjectionBorderManager(v33);
                        else
                          v34 = 0LL;
                        *((_QWORD *)this + 58) = v34;
                        if ( !*((_QWORD *)CDesktopManager::s_pDesktopManagerInstance + 58) )
                        {
                          v5 = -2147024882;
                          MilInstrumentationCheckHR_MaybeFailFast(0x14u, &dword_1800FD698, 1u, -2147024882, 0x1E0u, 0LL);
                          goto LABEL_103;
                        }
                        v35 = wil::out_param<wil::unique_any_t<wil::details::unique_storage<wil::details::resource_policy<HKEY__ *,long (*)(HKEY__ *),&long RegCloseKey(HKEY__ *),wistd::integral_constant<unsigned __int64,0>,HKEY__ *,HKEY__ *,0,std::nullptr_t>>>>(
                                Handles,
                                (char *)this + 208);
                        v5 = CCompositionEffectCache::Create((struct CCompositionEffectCache **)(v35 + 8));
                        wil::details::out_param_t<std::unique_ptr<CCompositionEffectCache>>::~out_param_t<std::unique_ptr<CCompositionEffectCache>>(Handles);
                        if ( v5 < 0 )
                        {
                          phkResulta = 482;
LABEL_79:
                          v37 = v5;
LABEL_102:
                          MilInstrumentationCheckHR_MaybeFailFast(0x14u, &dword_1800FD698, 1u, v37, phkResulta, 0LL);
LABEL_103:
                          CDesktopManager::NotifyRedirectionShutdown(this);
                          Data = &CDesktopManager::s_csDwmInstance;
                          LeaveCriticalSection(&CDesktopManager::s_csDwmInstance);
                          DwmRedirectionManagerShutdown();
                          EnterCriticalSection(&CDesktopManager::s_csDwmInstance);
                          if ( !EventW )
                            goto LABEL_110;
                          goto LABEL_104;
                        }
                        Factory = CWindowList::Initialize(*((CWindowList **)this + 53));
                        v5 = Factory;
                        if ( Factory < 0 )
                        {
                          phkResulta = 484;
                        }
                        else
                        {
                          SetLastError(0);
                          EventW = CreateEventW(0LL, 1, 0, 0LL);
                          if ( !EventW )
                          {
                            LastError = GetLastError();
                            v5 = LastError;
                            if ( LastError > 0 )
                              v5 = (unsigned __int16)LastError | 0x80070000;
                            phkResulta = 497;
LABEL_77:
                            if ( v5 >= 0 )
                              v5 = -2003304445;
                            goto LABEL_79;
                          }
                          SetLastError(0);
                          Thread = CreateThread(
                                     0LL,
                                     0LL,
                                     CDesktopManager::DwmEventThreadProc,
                                     EventW,
                                     0,
                                     (LPDWORD)this + 280);
                          *((_QWORD *)this + 141) = Thread;
                          if ( !Thread )
                          {
                            v39 = GetLastError();
                            v5 = v39;
                            if ( v39 > 0 )
                              v5 = (unsigned __int16)v39 | 0x80070000;
                            phkResulta = 506;
                            goto LABEL_77;
                          }
                          SetThreadDescription(Thread, L"uDWM Event Thread");
                          Handles[0] = EventW;
                          Handles[1] = *((HANDLE *)this + 141);
                          SetLastError(0);
                          if ( WaitForMultipleObjects(2u, Handles, 0, 0xFFFFFFFF) )
                          {
                            v40 = GetLastError();
                            v5 = v40;
                            if ( v40 > 0 )
                              v5 = (unsigned __int16)v40 | 0x80070000;
                            phkResulta = 523;
                            goto LABEL_77;
                          }
                          Factory = CDesktopManager::_InitializeWnf(this);
                          v5 = Factory;
                          if ( Factory >= 0 )
                          {
                            CDesktopManager::ReadProductType(this);
                            *((_BYTE *)this + 25) = 1;
LABEL_104:
                            CloseHandle(EventW);
                            goto LABEL_110;
                          }
                          phkResulta = 528;
                        }
                      }
                    }
                  }
                }
              }
            }
          }
        }
        v37 = Factory;
        goto LABEL_102;
      }
      phkResult = 422;
    }
LABEL_108:
    v13 = v4;
    goto LABEL_109;
  }
  phkResult = 393;
LABEL_19:
  v13 = -2147024882;
  v5 = -2147024882;
LABEL_109:
  MilInstrumentationCheckHR_MaybeFailFast(0x14u, &dword_1800FD698, 1u, v13, phkResult, 0LL);
LABEL_110:
  if ( v46 )
  {
    (*(void (__fastcall **)(__int64))(*(_QWORD *)v46 + 16LL))(v46);
    v46 = 0LL;
  }
  wil::details::unique_storage<wil::details::resource_policy<HKEY__ *,long (*)(HKEY__ *),&long RegCloseKey(HKEY__ *),wistd::integral_constant<unsigned __int64,0>,HKEY__ *,HKEY__ *,0,std::nullptr_t>>::~unique_storage<wil::details::resource_policy<HKEY__ *,long (*)(HKEY__ *),&long RegCloseKey(HKEY__ *),wistd::integral_constant<unsigned __int64,0>,HKEY__ *,HKEY__ *,0,std::nullptr_t>>(&hKey);
  return (unsigned int)v5;
}

// Hidden C++ exception states: #wind=6
CAnalogCompositorManager *__fastcall CAnalogCompositorManager::CAnalogCompositorManager(CAnalogCompositorManager *this)
{
  int v1; // eax
  __int64 v2; // rdx
  __int64 v3; // rax
  int AttachWatcher; // eax
  __int64 v5; // rdx
  __int64 v6; // rax
  int DetachWatcher; // eax
  int v9[34]; // [rsp+20h] [rbp-88h] BYREF
  wil::details::in1diag3 *retaddr; // [rsp+A8h] [rbp+0h]
  CAnalogCompositorManager *v11; // [rsp+B0h] [rbp+8h] BYREF
  void *v12; // [rsp+B8h] [rbp+10h]

  v11 = this;
  v12 = &qword_18011ECD0;
  qword_18011ECD0 = 0LL;
  Windows::Mirage::HolographicDriverDetectedWatcher::HolographicDriverDetectedWatcher((Windows::Mirage::HolographicDriverDetectedWatcher *)&unk_18011ECD8);
  qword_18011ED00 = 0LL;
  qword_18011ED08 = 0LL;
  qword_18011ED10 = 0LL;
  byte_18011ED18 = 0;
  qword_18011ED20 = 0LL;
  qword_18011ED28 = 0LL;
  qword_18011ED30 = 0LL;
  qword_18011ED38 = 0LL;
  qword_18011ED40 = 0LL;
  qword_18011ED48 = 0LL;
  qword_18011ED50 = 0LL;
  dword_18011ED58 = 0;
  byte_18011ED5C = 0;
  word_18011ED5D = 0;
  byte_18011ED5F = 0;
  qword_18011ED60 = 0LL;
  LODWORD(v11) = 0;
  if ( (*(int (__fastcall **)(_QWORD, const wchar_t *, CAnalogCompositorManager **))(**((_QWORD **)CDesktopManager::s_pDesktopManagerInstance
                                                                                      + 9)
                                                                                   + 8LL))(
         *((_QWORD *)CDesktopManager::s_pDesktopManagerInstance + 9),
         L"DisableHologramCompositor",
         &v11) < 0
    || !(_DWORD)v11 )
  {
    v1 = Windows::Mirage::HolographicDriverDetectedWatcher::RegisterForCMNotifications((Windows::Mirage::HolographicDriverDetectedWatcher *)&unk_18011ECD8);
    if ( v1 < 0 )
      wil::details::in1diag3::_FailFast_Hr(
        retaddr,
        (void *)0x26,
        (unsigned int)"clientcore\\windows\\dwm\\udwm\\analogcompositormanager.cpp",
        (const char *)(unsigned int)v1,
        v9[0]);
    LOBYTE(v2) = 0;
    v3 = wistd::function_void___cdecl_void__::function_void___cdecl_void____lambda_1fe009015b5481886de644cd00cd9360__void_(
           v9,
           v2);
    AttachWatcher = Windows::Mirage::HolographicDriverDetectedWatcher::CreateAttachWatcher(&unk_18011ECD8, v3);
    if ( AttachWatcher < 0 )
      wil::details::in1diag3::_FailFast_Hr(
        retaddr,
        (void *)0x2C,
        (unsigned int)"clientcore\\windows\\dwm\\udwm\\analogcompositormanager.cpp",
        (const char *)(unsigned int)AttachWatcher,
        v9[0]);
    LOBYTE(v5) = 0;
    v6 = wistd::function_void___cdecl_void__::function_void___cdecl_void____lambda_68ab246ca29dbf1f5c5163cf5c63f8ba__void_(
           v9,
           v5);
    DetachWatcher = Windows::Mirage::HolographicDriverDetectedWatcher::CreateDetachWatcher(&unk_18011ECD8, v6);
    if ( DetachWatcher < 0 )
      wil::details::in1diag3::_FailFast_Hr(
        retaddr,
        (void *)0x32,
        (unsigned int)"clientcore\\windows\\dwm\\udwm\\analogcompositormanager.cpp",
        (const char *)(unsigned int)DetachWatcher,
        v9[0]);
  }
  return (CAnalogCompositorManager *)&qword_18011ECD0;
}