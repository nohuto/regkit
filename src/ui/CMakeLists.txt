regkit_add_library(regkit_ui
    autocomplete.cpp
    dialog_layout.cpp
    feedback.cpp
    presets.cpp
)
target_include_directories(regkit_ui PRIVATE
    "${PROJECT_SOURCE_DIR}/resources"
)
target_link_libraries(regkit_ui PRIVATE
    regkit_win32
    comctl32
    dwmapi
    uxtheme
)
