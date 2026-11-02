regkit_add_library(regkit_records
    escaped_fields.cpp
)
target_link_libraries(regkit_records PRIVATE
    regkit_win32
)
