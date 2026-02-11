# CMakeLists.txt
# ==============================================================================
# Start
# ==============================================================================
cmake_minimum_required(VERSION 3.15)
project(xls2img VERSION 1.0.0 LANGUAGES C)

set(CMAKE_C_STANDARD 11)
set(CMAKE_C_STANDARD_REQUIRED ON)
set(CMAKE_C_EXTENSIONS OFF)

# all library files output to lib/ directory
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)
set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)

set(XLS2IMG_SOURCES
    src/xls2img_reader.c
    src/xls2img_images.c
)

# Define Windows version macros for compatibility
if(WIN32)
    add_compile_definitions(_WIN32_WINNT=0x0601)
    add_compile_definitions(WINVER=0x0601)
endif()

# ==============================================================================
# Library Definition
# ==============================================================================

# build dll
add_library(xls2img SHARED ${XLS2IMG_SOURCES})
set_target_properties(xls2img PROPERTIES
    OUTPUT_NAME "xls2img"
    DEBUG_POSTFIX "d"
    VERSION ${PROJECT_VERSION}
    SOVERSION ${PROJECT_VERSION_MAJOR}
)

# windows DLL export
if(WIN32)
    target_compile_definitions(xls2img PRIVATE XLS2IMG_BUILD_DLL)
endif()

# include
target_include_directories(xls2img PUBLIC ${CMAKE_CURRENT_SOURCE_DIR}/include)

target_compile_options(xls2img PRIVATE
    $<$<C_COMPILER_ID:MSVC>:/W3>
)

# ==============================================================================
# Tool Program Definition
# ==============================================================================

# c tool program (CLI tool, links shared lib, stays in lib/)
add_executable(xls2img_tool tool/xls2img_tool.c)
set_target_properties(xls2img_tool PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/tool  
    OUTPUT_NAME "xls2img_tool"
)
target_link_libraries(xls2img_tool xls2img)

# ==============================================================================
# Example Programs Definition
# ==============================================================================

# c example program
add_executable(xls2img_c_example examples/main.c)
set_target_properties(xls2img_c_example PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib
    OUTPUT_NAME "xls2img_c_example"
)
target_link_libraries(xls2img_c_example xls2img)

# cxx example program
enable_language(CXX)
add_executable(xls2img_cxx_example examples/main.cpp)
set_target_properties(xls2img_cxx_example PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib
    OUTPUT_NAME "xls2img_cxx_example"
)
target_link_libraries(xls2img_cxx_example xls2img)

# ==============================================================================
# Tool Distribution Setup
# ==============================================================================

# Create the 'tool' directory
file(MAKE_DIRECTORY ${CMAKE_BINARY_DIR}/tool)

# Custom commands to copy necessary files to the 'tool' directory after build
add_custom_command(TARGET xls2img POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
        "$<TARGET_FILE:xls2img>"
        "${CMAKE_BINARY_DIR}/tool/"
    COMMENT "Copying DLL to tool directory"
)

add_custom_command(TARGET xls2img_tool POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
        "$<TARGET_FILE:xls2img_tool>"
        "${CMAKE_BINARY_DIR}/tool/"
    COMMENT "Copying tool executable to tool directory"
)

# Copy the test file to the 'tool' directory
configure_file(
    "${CMAKE_CURRENT_SOURCE_DIR}/examples/test.xls"
    "${CMAKE_BINARY_DIR}/tool/test.xls"
    COPYONLY
)

# also copy the test file to the 'lib' directory
configure_file(
    "${CMAKE_CURRENT_SOURCE_DIR}/examples/test.xls"
    "${CMAKE_BINARY_DIR}/lib/test.xls"
    COPYONLY
)

# ==============================================================================
# Installation Rules
# ==============================================================================

install(TARGETS xls2img
    RUNTIME DESTINATION lib
    LIBRARY DESTINATION lib
    ARCHIVE DESTINATION lib
)

install(FILES ${CMAKE_CURRENT_SOURCE_DIR}/include/xls2img.h
    DESTINATION include
)

# ==============================================================================
# End
# ==============================================================================