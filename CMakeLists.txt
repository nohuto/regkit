cmake_minimum_required(VERSION 3.30)
project(RegKit VERSION 1.1.5 LANGUAGES C CXX)

set(CMAKE_CXX_SCAN_FOR_MODULES OFF)
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")
set(CMAKE_MSVC_DEBUG_INFORMATION_FORMAT ProgramDatabase)

include(FetchContent)
FetchContent_Declare(win32ui
    GIT_REPOSITORY https://github.com/nohuto/win32-ui.git
    GIT_TAG dc83962d3650d8328f5477a307dec3d2ed2b37f1
    EXCLUDE_FROM_ALL
)
FetchContent_MakeAvailable(win32ui)

win32ui_app_settings(regkit_compile_settings
    _WIN32_IE=0x0600
    REGKIT_VERSION_TEXT=L"${PROJECT_VERSION}"
)

include(cmake/RegKitTarget.cmake)
include(cmake/VerifyTargetGraph.cmake)

add_subdirectory(redist/pcre2)
add_subdirectory(src)

add_executable(RegKit WIN32
    src/launch/program.cpp
    resources/app.rc
)
regkit_configure_executable(RegKit)
target_include_directories(RegKit PRIVATE
    "${PROJECT_SOURCE_DIR}/resources"
)
target_link_libraries(RegKit PRIVATE
    regkit_frame
    regkit_cli
    regkit_dialogs
    regkit_ui
    regkit_win32
    comctl32
)
set_target_properties(RegKit PROPERTIES OUTPUT_NAME "regkit")

set_property(TARGET RegKit PROPERTY INTERPROCEDURAL_OPTIMIZATION_RELEASE TRUE)

if (CMAKE_SYSTEM_PROCESSOR MATCHES "^([Aa][Rr][Mm]64|aarch64)$")
    set(REGKIT_OFFREG_ARCH arm64)
elseif (CMAKE_SIZEOF_VOID_P EQUAL 4)
    set(REGKIT_OFFREG_ARCH x86)
else()
    set(REGKIT_OFFREG_ARCH x64)
endif()
set(REGKIT_OFFREG_DLL "${CMAKE_SOURCE_DIR}/redist/offreg/${REGKIT_OFFREG_ARCH}/offreg.dll")
if (NOT EXISTS "${REGKIT_OFFREG_DLL}")
    message(FATAL_ERROR "Missing ${REGKIT_OFFREG_DLL}. Offline Registry requires the Microsoft offreg.dll redistributable for this architecture.")
endif()

file(GLOB REGKIT_LANGUAGE_FILES CONFIGURE_DEPENDS "${CMAKE_SOURCE_DIR}/assets/lang/*.po")

add_custom_command(TARGET RegKit POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
        "${REGKIT_OFFREG_DLL}"
        "$<TARGET_FILE_DIR:RegKit>/offreg.dll"
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
        "${CMAKE_SOURCE_DIR}/redist/pcre2/LICENCE.md"
        "$<TARGET_FILE_DIR:RegKit>/licences/PCRE2-LICENCE.md"
    COMMAND ${CMAKE_COMMAND} -E make_directory "$<TARGET_FILE_DIR:RegKit>/assets/records"
    COMMAND ${CMAKE_COMMAND} -E rm -rf "$<TARGET_FILE_DIR:RegKit>/assets/icons"
    COMMAND ${CMAKE_COMMAND} -E make_directory "$<TARGET_FILE_DIR:RegKit>/assets/icons"
    COMMAND ${CMAKE_COMMAND} -E rm -rf "$<TARGET_FILE_DIR:RegKit>/assets/defaults"
    COMMAND ${CMAKE_COMMAND} -E make_directory "$<TARGET_FILE_DIR:RegKit>/assets/defaults"
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
        "${CMAKE_SOURCE_DIR}/assets/records/23H2.txt"
        "$<TARGET_FILE_DIR:RegKit>/assets/records/23H2.txt"
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
        "${CMAKE_SOURCE_DIR}/assets/records/24H2.txt"
        "$<TARGET_FILE_DIR:RegKit>/assets/records/24H2.txt"
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
        "${CMAKE_SOURCE_DIR}/assets/records/25H2.txt"
        "$<TARGET_FILE_DIR:RegKit>/assets/records/25H2.txt"
    COMMAND ${CMAKE_COMMAND} -E copy_directory
        "${CMAKE_SOURCE_DIR}/assets/icons"
        "$<TARGET_FILE_DIR:RegKit>/assets/icons"
    COMMAND ${CMAKE_COMMAND} -E copy_directory
        "${CMAKE_SOURCE_DIR}/assets/defaults"
        "$<TARGET_FILE_DIR:RegKit>/assets/defaults"
    COMMAND ${CMAKE_COMMAND} -E rm -rf "$<TARGET_FILE_DIR:RegKit>/assets/bitfields"
    COMMAND ${CMAKE_COMMAND} -E copy_directory
        "${CMAKE_SOURCE_DIR}/assets/bitfields"
        "$<TARGET_FILE_DIR:RegKit>/assets/bitfields"
    COMMAND ${CMAKE_COMMAND} -E rm -rf "$<TARGET_FILE_DIR:RegKit>/assets/lang"
    COMMAND ${CMAKE_COMMAND} -E make_directory "$<TARGET_FILE_DIR:RegKit>/assets/lang"
    COMMAND ${CMAKE_COMMAND} -E rm -rf "$<TARGET_FILE_DIR:RegKit>/assets/comments"
    COMMAND ${CMAKE_COMMAND} -E copy_directory
        "${CMAKE_SOURCE_DIR}/assets/comments"
        "$<TARGET_FILE_DIR:RegKit>/assets/comments"
)

if (REGKIT_LANGUAGE_FILES)
    add_custom_command(TARGET RegKit POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy_if_different ${REGKIT_LANGUAGE_FILES} "$<TARGET_FILE_DIR:RegKit>/assets/lang"
    )
endif()

if (MSVC)
    target_link_options(RegKit PRIVATE "/MANIFEST:NO" "/DEPENDENTLOADFLAG:0x800")
endif()

regkit_verify_dependency_order(
    regkit_win32
    regkit_records
    regkit_work
    regkit_ui
    regkit_registry
    regkit_regfile
    regkit_search
    regkit_trace
    regkit_defaults
    regkit_changes
    regkit_workspace
    regkit_browse
    regkit_dialogs
    regkit_frame
    regkit_cli
    RegKit
)

regkit_verify_include_order(win32 records work ui registry regfile search trace defaults changes workspace browse dialogs frame cli launch)
