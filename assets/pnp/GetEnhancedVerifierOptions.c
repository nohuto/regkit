void __fastcall GetEnhancedVerifierOptions(_FX_DRIVER_GLOBALS *FxDriverGlobals)
{
  _DRIVER_OBJECT *m_DriverObject; // rcx
  unsigned int value; // [rsp+30h] [rbp-59h] BYREF
  FxAutoRegKey hWdf; // [rsp+38h] [rbp-51h] BYREF
  FxAutoRegKey hKey; // [rsp+40h] [rbp-49h] BYREF
  _UNICODE_STRING parametersPath; // [rsp+48h] [rbp-41h] BYREF
  _UNICODE_STRING valueName; // [rsp+58h] [rbp-31h] BYREF
  _OBJECT_ATTRIBUTES ObjectAttributes; // [rsp+68h] [rbp-21h] BYREF
  wchar_t parametersPath_buffer[4]; // [rsp+98h] [rbp+Fh] BYREF
  wchar_t valueName_buffer[24]; // [rsp+A0h] [rbp+17h] BYREF

  value = 0;
  hKey.m_Key = 0LL;
  m_DriverObject = FxDriverGlobals->DriverObject.m_DriverObject;
  hWdf.m_Key = 0LL;
  wcscpy(parametersPath_buffer, L"Wdf");
  *(_QWORD *)&parametersPath.Length = 524294LL;
  parametersPath.Buffer = parametersPath_buffer;
  *(_QWORD *)&valueName.Length = 3145774LL;
  wcscpy(valueName_buffer, L"EnhancedVerifierOptions");
  valueName.Buffer = valueName_buffer;
  if ( (int)IoOpenDriverRegistryKey(m_DriverObject, 0LL, 131097LL, 0LL, &hWdf) >= 0 )
  {
    *(&ObjectAttributes.Length + 1) = 0;
    memset(&ObjectAttributes.Attributes + 1, 0, 20);
    ObjectAttributes.RootDirectory = hWdf.m_Key;
    ObjectAttributes.Length = 48;
    ObjectAttributes.ObjectName = &parametersPath;
    ObjectAttributes.Attributes = 576;
    if ( ZwOpenKey(&hKey.m_Key, 0x20019u, &ObjectAttributes) >= 0
      && FxRegKey::_QueryULong(hKey.m_Key, &valueName, &value) >= 0 )
    {
      FxDriverGlobals->FxEnhancedVerifierOptions = value;
    }
  }
  if ( hWdf.m_Key )
    ZwClose(hWdf.m_Key);
  if ( hKey.m_Key )
    ZwClose(hKey.m_Key);
}