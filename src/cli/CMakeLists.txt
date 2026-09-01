regkit_add_library(regkit_cli
    reg_command.cpp
)
target_include_directories(regkit_cli PUBLIC
    "${PROJECT_SOURCE_DIR}/src"
)
target_link_libraries(regkit_cli PRIVATE
    regkit_registry_io
    regkit_regfile
    regkit_registry
    regkit_win32
    advapi32
)
