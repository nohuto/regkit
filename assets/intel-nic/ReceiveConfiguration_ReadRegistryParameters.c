void __fastcall ReceiveConfiguration::ReadRegistryParameters(struct ADAPTER_CONTEXT **this)
{
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)(this + 11),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"*ReceiveBuffers",
    0x80u,
    0x1000u,
    0x200u,
    1,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(this + 3),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"ReceiveBuffersOverride",
    0,
    1u,
    1u,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)(this + 8),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"MaxPacketCountPerDPC",
    8u,
    0xFFFFu,
    0x100u,
    1,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)((char *)this + 76),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"MaxPacketCountPerIndicate",
    1u,
    0xFFFFu,
    0x40u,
    1,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)((char *)this + 100),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"RxDescriptorCountPerTailWrite",
    4u,
    0x1000u,
    8u,
    1,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)((char *)this + 52),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"MinHardwareOwnedPacketCount",
    8u,
    0x1000u,
    0x20u,
    1,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)(this + 5),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"RxBufferPad",
    0,
    0x3Fu,
    0xAu,
    1,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(this + 4),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"RegForceRxPathSerialization",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(this + 2),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"EnableRxDescriptorChaining",
    0,
    1u,
    1u,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)(this + 14),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"EnableAdaptiveQueuing",
    0,
    1u,
    1u,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)((char *)this + 124),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"AdaptiveQSize",
    0x40u,
    0x2000u,
    0x80u,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)(this + 17),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"AdaptiveQWorkSet",
    0x20u,
    0x2000u,
    0x60u,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)((char *)this + 148),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"AdaptiveQHysteresis",
    0x10u,
    0x400u,
    0x40u,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(this + 21),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"PadReceiveBuffer",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(this + 20),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"StoreBadPackets",
    0,
    1u,
    0,
    0,
    0);
}
