void __fastcall HeaderSplitConfiguration::ReadRegistryParameters(struct ADAPTER_CONTEXT **this)
{
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)(this + 7),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"*HeaderDataSplit",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)((char *)this + 68),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"HDSplitSize",
    0x80u,
    0x3C0u,
    0x80u,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(this + 10),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"HDSplitAlways",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)(this + 11),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"HDSplitLocation",
    0,
    3u,
    2u,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)((char *)this + 100),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"HDSplitBufferPad",
    0,
    2u,
    2u,
    0,
    0);
}
