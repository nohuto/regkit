regkit_add_library(regkit_work
    key_watcher.cpp
)
target_link_libraries(regkit_work PRIVATE
    regkit_win32
)
