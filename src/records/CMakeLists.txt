regkit_add_library(regkit_records
    escaped_fields.cpp
)
target_include_directories(regkit_records PUBLIC
    "${PROJECT_SOURCE_DIR}/src"
)
