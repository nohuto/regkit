regkit_add_library(regkit_defaults
    default_data.cpp
    default_loader.cpp
)
target_include_directories(regkit_defaults PUBLIC
    "${PROJECT_SOURCE_DIR}/src"
)
target_link_libraries(regkit_defaults PRIVATE
    regkit_regfile
    regkit_registry
)
