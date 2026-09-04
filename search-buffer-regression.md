# Search buffer regression and correction

## Outcome

The rewritten search engine is slower because its reusable registry-enumeration buffers retain the largest logical size encountered by each worker.

The buffer allocation should be retained, but the logical size passed to the registry API must be reset for every key. The current implementation retains both, causing later `RegEnumValueW` calls to receive an unnecessarily large buffer length.

This defect exists in the live and offline registry backends. It accounts for the measured search regression from approximately 10 seconds to approximately 30 seconds.

Do not remove reusable worker buffers and do not increase the worker count to hide the problem.

## Measured evidence

The engine was measured directly in an x64 Release build. UI painting, result-list updates, and dialog handling were excluded from the measurement.

The query was `networkthrottlingindex` with value-name and value-data searching enabled.

| Engine state | Scope | Completion time |
|---|---|---:|
| Previous committed engine | Five standard roots | 7.8 to 8.0 seconds |
| Current rewritten engine | Five standard roots | 19.8 to 28.3 seconds |
| Previous committed engine | `HKEY_LOCAL_MACHINE` | 3.9 to 4.2 seconds |
| Current rewritten engine | `HKEY_LOCAL_MACHINE` | 7.6 to 10.1 seconds |
| Current engine with scratch reuse disabled | `HKEY_LOCAL_MACHINE` | 3.8 to 4.0 seconds |
| Current engine with corrected scratch sizing | `HKEY_LOCAL_MACHINE` | 3.6 to 4.0 seconds |
| Corrected scratch sizing with the current four workers | Five standard roots | 7.5 to 8.3 seconds |

The corrected four-worker engine matches the previous engine. This isolates the regression to buffer sizing and shows that eight workers are not required to recover the lost performance.

## Exact cause

The search worker creates one `EnumerationScratch` in `src/search/search.cpp` and passes it to every key enumerated by that worker:

```cpp
EnumerationScratch scratch;
```

The call to `RegistryStore::EnumKeyStreaming` correctly passes the same scratch object:

```cpp
enum_max_data, &scratch, false
```

This ownership is correct and should remain.

The defect is inside `src/registry/live_registry.cpp`:

```cpp
if (data.size() < needed) {
  data.resize(needed);
}
```

That code only increases `data.size()`. If one key requires a large data buffer, the vector keeps that large logical size for every later key handled by the worker.

The next enumeration initializes the native API length from the retained size:

```cpp
DWORD data_length =
    include_data ? static_cast<DWORD>(data.size()) : 0;
```

`RegEnumValueW` therefore receives the largest length previously encountered by that worker instead of the current key's reported maximum data length.

The same pattern affects the value-name and subkey-name buffers. Their impact is smaller because registry names are normally much shorter than value data, but they should follow the same correct ownership rule.

The offline backend in `src/registry/offline_registry.cpp` contains the same implementation error.

## Required change in the live backend

Open `src/registry/live_registry.cpp` and find the buffer preparation inside `EnumKeyStreaming`.

Replace:

```cpp
if (name.size() < static_cast<size_t>(max_value_name_length) + 1) {
  name.resize(static_cast<size_t>(max_value_name_length) + 1);
}
if (include_data) {
  const size_t needed = std::min(max_value_data_length, max_data_size);
  if (data.size() < needed) {
    data.resize(needed);
  }
}
```

with:

```cpp
name.resize(static_cast<size_t>(max_value_name_length) + 1);
if (include_data) {
  const size_t needed = std::min(max_value_data_length, max_data_size);
  data.resize(needed);
}
```

Then replace the subkey-name preparation:

```cpp
if (name.size() < static_cast<size_t>(max_subkey_length) + 1) {
  name.resize(static_cast<size_t>(max_subkey_length) + 1);
}
```

with:

```cpp
name.resize(static_cast<size_t>(max_subkey_length) + 1);
```

Calling `resize` with a smaller value changes the logical size without reducing vector capacity. Calling it with a larger value reuses existing capacity whenever possible and allocates only when the retained capacity is insufficient.

The worker therefore keeps the allocation benefit while `RegEnumValueW` and `RegEnumKeyExW` receive lengths appropriate for the current key.

## Required change in the offline backend

Apply the identical changes inside `EnumKeyStreaming` in `src/registry/offline_registry.cpp`.

Replace the grow-only preparation for:

- `buffers.value_name`
- `buffers.value_data`
- `buffers.subkey_name`

with unconditional `resize` calls using the current key's values from `query_info`.

The target code should be:

```cpp
std::wstring& name = buffers.value_name;
std::vector<BYTE>& data = buffers.value_data;
name.resize(static_cast<size_t>(max_value_name_length) + 1);
if (include_data) {
  const size_t needed = std::min(max_value_data_length, max_data_size);
  data.resize(needed);
}
```

The subkey buffer should use:

```cpp
std::wstring& name = buffers.subkey_name;
name.resize(static_cast<size_t>(max_subkey_length) + 1);
```

The live and offline implementations must remain structurally identical here. Allowing one backend to keep the grow-only behavior would make the same regression reappear for offline hive searches.

## Progress-throttling correction

There is a separate implementation mistake in `src/search/search.cpp`.

After every private node chunk, the worker currently calls:

```cpp
flush_progress(true);
```

The `true` argument forces publication and bypasses both intended thresholds:

- `kProgressKeyInterval`, currently 256 keys
- `kProgressTickInterval`, currently 200 ms

With a 24-key chunk, a 448,000-key scan can attempt approximately 18,700 forced progress publications.

Change the call immediately after `publish_batch(batch)` and before the `active` counter is decremented to:

```cpp
flush_progress(false);
```

Keep the final call at the end of the worker as:

```cpp
flush_progress(true);
```

The non-forced chunk-boundary call publishes only when a threshold has been reached. The final forced call guarantees that the last partial count is not lost.

This progress issue was not responsible for the threefold traversal regression in the direct measurements, but correcting it removes unnecessary callback and atomic work from the search loop.

## Code that must remain

Keep these parts of the rewrite:

- One `EnumerationScratch` owned by each search worker
- Passing `&scratch` to `RegistryStore::EnumKeyStreaming`
- The four-worker local provider limit
- Worker-owned result batches
- Direct string-pointer and cached-data use
- Deferred hydration for a value-name match
- The compact search-result structure
- Bounded pending-result batches
- One final forced progress publication

These mechanisms are not responsible for the regression.

## Changes that must not be used

Do not apply any of the following:

- Removing scratch reuse and returning to allocations for every key
- Clearing or swapping the buffers after every key
- Calling `shrink_to_fit`
- Raising the local worker limit from four to eight as the correction
- Keeping the maximum logical buffer size and passing a separate guessed length
- Adding a hard maximum registry-value size unless the user selected a maximum-size filter
- Disabling value-data searching to make this query appear faster

Removing scratch reuse happens to restore the previous timing, but it discards the intended allocation optimization. Correct logical sizing retains both speed and minimal allocation.

Increasing the worker count also reduces the observed delay, but it only hides the defect by performing more oversized native calls concurrently.

## Final expected path

For each key, the worker should perform this sequence:

1. Query the current key's maximum name and data lengths
2. Resize the reusable buffers to those logical lengths
3. Retain any larger allocated capacity from earlier keys
4. Pass the current logical lengths to the native enumeration functions
5. Enumerate values and children
6. Reuse the same storage for the next key

After these changes, a sparse full-registry search should return to approximately the previous completion time while retaining the rewritten queue, result model, and allocation improvements.
