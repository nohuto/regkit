regkit_add_library(regkit_search
    compare.cpp
    replace.cpp
    result_file.cpp
    search.cpp
)
target_include_directories(regkit_search PUBLIC
    "${PROJECT_SOURCE_DIR}/src"
)
target_link_libraries(regkit_search PRIVATE
    regkit_regfile
    regkit_records
    regkit_registry
    regkit_win32
)
