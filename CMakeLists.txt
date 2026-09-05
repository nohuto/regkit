cmake_minimum_required(VERSION 3.23)
project(RegKit LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

option(REGKIT_REPORT_COMPILE_TIME "Print MSVC front-end and back-end timing" OFF)

message(STATUS "RegKit platform: '${CMAKE_GENERATOR_PLATFORM}' pointer size: ${CMAKE_SIZEOF_VOID_P}")

add_library(regkit_compile_settings INTERFACE)
target_compile_features(regkit_compile_settings INTERFACE cxx_std_20)
target_compile_definitions(regkit_compile_settings INTERFACE
    UNICODE
    _UNICODE
    NOMINMAX
    WIN32_LEAN_AND_MEAN
    _WIN32_IE=0x0600
)
if (MSVC)
    target_compile_options(regkit_compile_settings INTERFACE
        /EHsc
        /FS
        /W4
        /permissive-
    )
    if (REGKIT_REPORT_COMPILE_TIME)
        target_compile_options(regkit_compile_settings INTERFACE /Bt+)
    endif()
endif()

include(cmake/RegKitTarget.cmake)
include(cmake/VerifyTargetGraph.cmake)

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
    regkit_registry_io
    regkit_editors
    regkit_appearance
    regkit_win32
    advapi32
    comctl32
    shell32
)
set_target_properties(RegKit PROPERTIES OUTPUT_NAME "regkit")

include(CheckIPOSupported)
check_ipo_supported(RESULT REGKIT_IPO_SUPPORTED OUTPUT REGKIT_IPO_ERROR)
if (REGKIT_IPO_SUPPORTED)
    set_property(TARGET RegKit PROPERTY INTERPROCEDURAL_OPTIMIZATION_RELEASE TRUE)
endif()

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

add_custom_command(TARGET RegKit POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
        "${REGKIT_OFFREG_DLL}"
        "$<TARGET_FILE_DIR:RegKit>/offreg.dll"
    COMMAND ${CMAKE_COMMAND} -E make_directory "$<TARGET_FILE_DIR:RegKit>/assets/records"
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
)

if (MSVC)
    target_link_options(RegKit PRIVATE "/MANIFEST:NO")
endif()

regkit_verify_dependency_order(
    regkit_win32
    regkit_records
    regkit_work
    regkit_registry
    regkit_regfile
    regkit_search
    regkit_trace
    regkit_defaults
    regkit_changes
    regkit_browse
    regkit_workspace
    regkit_appearance
    regkit_editors
    regkit_registry_io
    regkit_search_dialogs
    regkit_trace_dialog
    regkit_security
    regkit_frame
    regkit_cli
    RegKit
)
