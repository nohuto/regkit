regkit_add_library(regkit_changes
    change_history.cpp
    key_snapshot.cpp
    undo_stack.cpp
    value_comments.cpp
)
target_include_directories(regkit_changes PUBLIC
    "${PROJECT_SOURCE_DIR}/src"
)
target_link_libraries(regkit_changes PRIVATE
    regkit_records
    regkit_registry
    regkit_win32
)
