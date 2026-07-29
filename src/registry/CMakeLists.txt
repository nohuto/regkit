regkit_add_library(regkit_registry
    live_registry.cpp
    offline_registry.cpp
    registry_path.cpp
    registry_provider.cpp
    value_format.cpp
    virtual_registry.cpp
)
target_include_directories(regkit_registry PUBLIC
    "${PROJECT_SOURCE_DIR}/src"
)
target_link_libraries(regkit_registry PRIVATE
    regkit_win32
    advapi32
    shlwapi
)
