void __fastcall ReceiveSideCoalescing::ReadRegistryParameters(struct ADAPTER_CONTEXT **this)
{
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)((char *)this + 36),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"*RSCIPv4",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)((char *)this + 44),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"*RSCIPv6",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(this + 2),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"ForceRscEnabled",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)(this + 3),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"RscMode",
    0,
    2u,
    1u,
    0,
    0);
}
