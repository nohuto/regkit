regkit_add_library(regkit_trace
    trace_data.cpp
    trace_loader.cpp
    trace_parser.cpp
)
target_include_directories(regkit_trace PUBLIC
    "${PROJECT_SOURCE_DIR}/src"
)
target_link_libraries(regkit_trace PRIVATE
    regkit_registry
    regkit_win32
)

regkit_add_library(regkit_trace_dialog
    trace_dialog.cpp
)
target_include_directories(regkit_trace_dialog PRIVATE
    "${PROJECT_SOURCE_DIR}/resources"
)
target_link_libraries(regkit_trace_dialog PRIVATE
    regkit_editors
    regkit_appearance
    regkit_registry
    regkit_trace
    regkit_win32
    comctl32
)
