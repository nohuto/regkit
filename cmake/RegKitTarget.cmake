function(regkit_add_library target)
    add_library(${target} STATIC ${ARGN})
    target_link_libraries(${target} PRIVATE regkit_compile_settings)
    target_include_directories(${target} PRIVATE
        "${PROJECT_SOURCE_DIR}/include"
    )
endfunction()

function(regkit_configure_executable target)
    target_link_libraries(${target} PRIVATE regkit_compile_settings)
    target_include_directories(${target} PRIVATE
        "${PROJECT_SOURCE_DIR}/include"
        "${PROJECT_SOURCE_DIR}/src"
    )
endfunction()
