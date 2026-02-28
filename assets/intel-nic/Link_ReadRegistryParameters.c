void __fastcall Link::ReadRegistryParameters(struct ADAPTER_CONTEXT **this)
{
  RegistryKey<enum HdSplitLocation>::Initialize(
    (struct ADAPTER_CONTEXT *)((char *)*this + 992),
    *this,
    *((NDIS_HANDLE *)*this + 383),
    (PUCHAR)"*FlowControl",
    0,
    4u,
    4u,
    0,
    1);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (struct ADAPTER_CONTEXT *)((char *)*this + 980),
    *this,
    *((NDIS_HANDLE *)*this + 383),
    (PUCHAR)"*SpeedDuplex",
    0,
    0xC350u,
    0,
    0,
    1);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (struct ADAPTER_CONTEXT *)((char *)*this + 1004),
    *this,
    *((NDIS_HANDLE *)*this + 383),
    (PUCHAR)"FecMode",
    0,
    3u,
    0,
    0,
    1);
  (**(void (__fastcall ***)(struct ADAPTER_CONTEXT *))this[1])(this[1]);
}
