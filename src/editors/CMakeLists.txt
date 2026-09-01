regkit_add_library(regkit_editors
    binary_editor.cpp
    binary_text.cpp
    comment_editor.cpp
    dialog_support.cpp
    export_dialog.cpp
    hive_dialog.cpp
    value_editor.cpp
)
target_include_directories(regkit_editors PRIVATE
    "${PROJECT_SOURCE_DIR}/resources"
)
target_include_directories(regkit_editors PUBLIC
    "${PROJECT_SOURCE_DIR}/src"
)
target_link_libraries(regkit_editors PRIVATE
    regkit_appearance
    regkit_registry
    regkit_win32
    comctl32
    shell32
    uxtheme
)
