regkit_add_library(regkit_work
    session.cpp
)
target_include_directories(regkit_work PUBLIC
    "${PROJECT_SOURCE_DIR}/src"
)
