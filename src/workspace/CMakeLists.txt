regkit_add_library(regkit_workspace
    dialog_state.cpp
    favorites.cpp
    settings.cpp
    tab_state.cpp
    tree_state.cpp
)
target_link_libraries(regkit_workspace PRIVATE
    regkit_records
    regkit_win32
)
