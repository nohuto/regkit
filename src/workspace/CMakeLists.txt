regkit_add_library(regkit_workspace
    favorites.cpp
    recent_items.cpp
    settings.cpp
    tab_state.cpp
    tree_state.cpp
)
target_include_directories(regkit_workspace PUBLIC
    "${PROJECT_SOURCE_DIR}/src"
)
target_link_libraries(regkit_workspace PRIVATE
    regkit_records
    regkit_win32
    advapi32
    ole32
    shell32
)
