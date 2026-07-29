regkit_add_library(regkit_browse
    registry_tree.cpp
    value_list.cpp
)
target_link_libraries(regkit_browse PRIVATE
    regkit_registry
    regkit_work
    regkit_win32
    comctl32
)

regkit_add_library(regkit_appearance
    font_dialog.cpp
    theme.cpp
    theme_presets.cpp
    ../win32/icon_resources.cpp
)
target_link_libraries(regkit_appearance PRIVATE
    regkit_win32
    advapi32
    comctl32
    comdlg32
    dwmapi
    uxtheme
)

regkit_add_library(regkit_editors
    ui_helpers.cpp
    value_dialogs.cpp
)
target_include_directories(regkit_editors PRIVATE
    "${PROJECT_SOURCE_DIR}/resources"
)
target_link_libraries(regkit_editors PRIVATE
    regkit_appearance
    regkit_registry
    regkit_win32
    comctl32
    comdlg32
    shell32
    uxtheme
)

regkit_add_library(regkit_registry_io
    registry_io.cpp
)
target_link_libraries(regkit_registry_io PRIVATE
    regkit_editors
    regkit_regfile
    regkit_win32
    comdlg32
    shell32
)

regkit_add_library(regkit_search_dialogs
    replace_dialog.cpp
    search_dialog.cpp
)
target_link_libraries(regkit_search_dialogs PRIVATE
    regkit_browse
    regkit_editors
    regkit_search
    regkit_registry
    regkit_appearance
    regkit_win32
    comctl32
    shell32
    uxtheme
)

regkit_add_library(regkit_trace_dialog
    trace_dialog.cpp
)
target_link_libraries(regkit_trace_dialog PRIVATE
    regkit_editors
    regkit_appearance
    regkit_registry
    regkit_trace
    regkit_win32
    comctl32
)

regkit_add_library(regkit_security
    registry_security.cpp
)
target_link_libraries(regkit_security PRIVATE
    regkit_registry
    aclui
    advapi32
)

regkit_add_library(regkit_frame
    app_window.cpp
    app_window_commands.cpp
    theme_presets_dialog.cpp
    toolbar.cpp
)
target_include_directories(regkit_frame PRIVATE
    "${PROJECT_SOURCE_DIR}/resources"
)
target_link_libraries(regkit_frame PRIVATE
    regkit_browse
    regkit_changes
    regkit_workspace
    regkit_appearance
    regkit_editors
    regkit_registry_io
    regkit_search_dialogs
    regkit_trace_dialog
    regkit_security
    regkit_search
    regkit_trace
    regkit_defaults
    regkit_regfile
    regkit_registry
    regkit_win32
    advapi32
    comctl32
    comdlg32
    ole32
    pathcch
    shell32
    shlwapi
    uxtheme
)
