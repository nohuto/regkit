regkit_add_library(regkit_win32
    file_dialog.cpp
    file_text.cpp
    handle_owner.cpp
    process_rights.cpp
    registry_native.cpp
    restart.cpp
    shell_paths.cpp
    system_error.cpp
    text_transform.cpp
    window_metrics.cpp
)
target_include_directories(regkit_win32 PUBLIC
    "${PROJECT_SOURCE_DIR}/src"
)
target_link_libraries(regkit_win32 PRIVATE
    advapi32
    ole32
    shell32
    userenv
    wtsapi32
)
