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

regkit_add_library(regkit_search_dialogs
    query_dialog.cpp
    query_prompts.cpp
    replace_dialog.cpp
)
target_link_libraries(regkit_search_dialogs PRIVATE
    regkit_browse
    regkit_editors
    regkit_search
    regkit_registry
    regkit_appearance
    regkit_win32
    comctl32
    shell32
    uxtheme
)
