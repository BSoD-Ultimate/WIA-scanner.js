cmake_minimum_required(VERSION 3.0.2) 
message(STATUS "CMake version: ${CMAKE_VERSION} " )

set(CMAKE_MODULE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/CMake;${CMAKE_MODULE_PATH}")
include(ExternalProject)

project(wia-scanner-js)

message(STATUS "Current Platform: ${CMAKE_SYSTEM_NAME}")

# where to install PDB files
set(INSTALL_DEBUG_SYMBOLS ON CACHE BOOL "install generated pdb files")
set(DEBUG_SYMBOLS_OUTPUT_PATH "${CMAKE_INSTALL_PREFIX}/.debug" CACHE STRING "debug symbols output path")

add_subdirectory(src)


# Installation
set(INSTALL_FILE_LIST
  "$<TARGET_LINKER_FILE:wia-scanner-js>"
  "$<TARGET_FILE:wia-scanner-js>"
  "${CMAKE_CURRENT_SOURCE_DIR}/src/index.js"
  "${CMAKE_CURRENT_SOURCE_DIR}/src/package.json"
)

install(FILES 
  ${INSTALL_FILE_LIST}
  DESTINATION 
  "${CMAKE_INSTALL_PREFIX}/wia-scanner-js"
  COMPONENT publish
)

if(INSTALL_DEBUG_SYMBOLS)
  # install debug symbols to the target directory
  set(DEBUG_SYMBOLS_LIST 
    "$<TARGET_PDB_FILE:wia-scanner-js>"
  )

  install(FILES 
    ${DEBUG_SYMBOLS_LIST}
    DESTINATION 
      "${DEBUG_SYMBOLS_OUTPUT_PATH}"
    COMPONENT publish
  )
endif()

