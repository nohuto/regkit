void __fastcall Power::ReadRegistryParameters(Power *this)
{
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(*((_QWORD *)this + 1) + 1196LL),
    *((struct ADAPTER_CONTEXT **)this + 1),
    *(NDIS_HANDLE *)(*((_QWORD *)this + 1) + 3064LL),
    (PUCHAR)"EnablePowerManagement",
    0,
    1u,
    1u,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(*((_QWORD *)this + 1) + 1188LL),
    *((struct ADAPTER_CONTEXT **)this + 1),
    *(NDIS_HANDLE *)(*((_QWORD *)this + 1) + 3064LL),
    (PUCHAR)"EnablePME",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)(*((_QWORD *)this + 1) + 1208LL),
    *((struct ADAPTER_CONTEXT **)this + 1),
    *(NDIS_HANDLE *)(*((_QWORD *)this + 1) + 3064LL),
    (PUCHAR)"WakeFromS5",
    0,
    0xFFFFu,
    2u,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(*((_QWORD *)this + 1) + 1180LL),
    *((struct ADAPTER_CONTEXT **)this + 1),
    *(NDIS_HANDLE *)(*((_QWORD *)this + 1) + 3064LL),
    (PUCHAR)"*DeviceSleepOnDisconnect",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(*((_QWORD *)this + 1) + 1220LL),
    *((struct ADAPTER_CONTEXT **)this + 1),
    *(NDIS_HANDLE *)(*((_QWORD *)this + 1) + 3064LL),
    (PUCHAR)"EnableModernStandby",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(*((_QWORD *)this + 1) + 1232LL),
    *((struct ADAPTER_CONTEXT **)this + 1),
    *(NDIS_HANDLE *)(*((_QWORD *)this + 1) + 3064LL),
    (PUCHAR)"*PMARPOffload",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(*((_QWORD *)this + 1) + 1240LL),
    *((struct ADAPTER_CONTEXT **)this + 1),
    *(NDIS_HANDLE *)(*((_QWORD *)this + 1) + 3064LL),
    (PUCHAR)"*PMNSOffload",
    0,
    1u,
    0,
    0,
    0);
}
