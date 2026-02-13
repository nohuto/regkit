cmake_minimum_required(VERSION 3.23)
project(RegKit LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

add_executable(RegKit WIN32
    src/app_entry.cpp
    src/app/app_window.cpp
    src/app/app_window_commands.cpp
    src/app/value_dialogs.cpp
    src/app/registry_tree.cpp
    src/app/value_list.cpp
    src/app/theme.cpp
    src/app/theme_presets.cpp
    src/app/theme_presets_dialog.cpp
    src/app/toolbar.cpp
    src/app/favorites_store.cpp
    src/app/font_dialog.cpp
    src/app/registry_io.cpp
    src/app/search_dialog.cpp
    src/app/trace_dialog.cpp
    src/app/replace_dialog.cpp
    src/app/registry_security.cpp
    src/app/ui_helpers.cpp
    src/registry/registry_provider.cpp
    src/registry/search_engine.cpp
    src/win32/win32_helpers.cpp
    src/win32/icon_resources.cpp
    resources/app.rc
)

target_include_directories(RegKit PRIVATE include resources)

set_target_properties(RegKit PROPERTIES OUTPUT_NAME "regkit")

target_compile_definitions(RegKit PRIVATE
    UNICODE
    _UNICODE
    NOMINMAX
    WIN32_LEAN_AND_MEAN
    _WIN32_IE=0x0600
)

target_link_libraries(RegKit PRIVATE
    comctl32
    comdlg32
    aclui
    advapi32
    dwmapi
    ole32
    pathcch
    shell32
    shlwapi
    userenv
    uxtheme
    wtsapi32
)

add_custom_command(TARGET RegKit POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory "$<TARGET_FILE_DIR:RegKit>/assets/traces"
    COMMAND ${CMAKE_COMMAND} -E make_directory "$<TARGET_FILE_DIR:RegKit>/assets/icons"
    COMMAND ${CMAKE_COMMAND} -E make_directory "$<TARGET_FILE_DIR:RegKit>/assets/defaults"
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
        "${CMAKE_SOURCE_DIR}/assets/associate_regkit_reg.bat"
        "$<TARGET_FILE_DIR:RegKit>/assets/associate_regkit_reg.bat"
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
        "${CMAKE_SOURCE_DIR}/assets/remove_regkit_reg.bat"
        "$<TARGET_FILE_DIR:RegKit>/assets/remove_regkit_reg.bat"
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
        "${CMAKE_SOURCE_DIR}/assets/traces/23H2.txt"
        "$<TARGET_FILE_DIR:RegKit>/assets/traces/23H2.txt"
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
        "${CMAKE_SOURCE_DIR}/assets/traces/24H2.txt"
        "$<TARGET_FILE_DIR:RegKit>/assets/traces/24H2.txt"
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
        "${CMAKE_SOURCE_DIR}/assets/traces/25H2.txt"
        "$<TARGET_FILE_DIR:RegKit>/assets/traces/25H2.txt"
    COMMAND ${CMAKE_COMMAND} -E copy_directory
        "${CMAKE_SOURCE_DIR}/assets/icons"
        "$<TARGET_FILE_DIR:RegKit>/assets/icons"
    COMMAND ${CMAKE_COMMAND} -E copy_directory
        "${CMAKE_SOURCE_DIR}/assets/defaults"
        "$<TARGET_FILE_DIR:RegKit>/assets/defaults"
)

if (MSVC)
    target_link_options(RegKit PRIVATE "/MANIFEST:NO")
endif()
