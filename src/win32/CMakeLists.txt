regkit_add_library(regkit_win32
    file_dialog.cpp
    file_text.cpp
    key_handle_snapshot.cpp
    process_rights.cpp
    registry_native.cpp
    restart.cpp
    shell_integration.cpp
    shell_paths.cpp
    translation.cpp
)
target_link_libraries(regkit_win32 PUBLIC win32ui::win32ui)
