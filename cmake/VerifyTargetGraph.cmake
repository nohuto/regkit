function(regkit_verify_dependency_order)
    set(allowed_targets regkit_compile_settings)

    foreach(target IN LISTS ARGN)
        if (NOT TARGET ${target})
            message(FATAL_ERROR "Unknown RegKit target in dependency order: ${target}")
        endif()

        get_target_property(dependencies ${target} LINK_LIBRARIES)
        if (dependencies STREQUAL "dependencies-NOTFOUND")
            set(dependencies)
        endif()

        foreach(dependency IN LISTS dependencies)
            if (TARGET ${dependency})
                list(FIND allowed_targets ${dependency} dependency_index)
                if (dependency_index EQUAL -1)
                    message(FATAL_ERROR
                        "Invalid RegKit dependency: ${target} -> ${dependency}. "
                        "Dependencies must point to an earlier target.")
                endif()
            endif()
        endforeach()

        list(APPEND allowed_targets ${target})
    endforeach()
endfunction()
