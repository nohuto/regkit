regkit_add_library(regkit_win32
    file_text.cpp
    process_rights.cpp
    registry_native.cpp
    shell_paths.cpp
    win32_helpers.cpp
)
target_include_directories(regkit_win32 PUBLIC
    "${PROJECT_SOURCE_DIR}/src"
)
target_link_libraries(regkit_win32 PRIVATE
    advapi32
    ole32
    pathcch
    shell32
    userenv
    wtsapi32
)
