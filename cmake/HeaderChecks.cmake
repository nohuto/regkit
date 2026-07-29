function(regkit_add_header_checks target)
    set(check_sources)

    foreach(header IN LISTS ARGN)
        string(MAKE_C_IDENTIFIER "${header}" header_id)
        set(check_source "${CMAKE_CURRENT_BINARY_DIR}/header_checks/${header_id}.cpp")
        file(GENERATE
            OUTPUT "${check_source}"
            CONTENT "#include \"${header}\"\n"
        )
        set_source_files_properties("${check_source}" PROPERTIES GENERATED TRUE)
        list(APPEND check_sources "${check_source}")
    endforeach()

    add_library(${target} OBJECT ${check_sources})
    target_link_libraries(${target} PRIVATE regkit_compile_settings)
    target_include_directories(${target} PRIVATE
        "${PROJECT_SOURCE_DIR}/include"
        "${PROJECT_SOURCE_DIR}/src"
    )
endfunction()
