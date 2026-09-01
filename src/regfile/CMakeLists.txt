regkit_add_library(regkit_regfile
    reg_file.cpp
)
target_include_directories(regkit_regfile PUBLIC
    "${PROJECT_SOURCE_DIR}/src"
)
target_link_libraries(regkit_regfile PRIVATE
    regkit_registry
    regkit_win32
)

regkit_add_library(regkit_registry_io
    registry_transfer.cpp
)
target_link_libraries(regkit_registry_io PRIVATE
    regkit_editors
    regkit_regfile
    regkit_win32
    shell32
)

