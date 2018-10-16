cmake_minimum_required(VERSION 3.2.0) 

message(STATUS "CMake version: ${CMAKE_VERSION}")
message(STATUS "Current Platform is ${${CMAKE_SYSTEM_NAME}}")
message(STATUS "Current Generator is ${CMAKE_GENERATOR}")

set(CMAKE_MODULE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/CMake;${CMAKE_MODULE_PATH}")
include(PrecompiledHeader)

# determines which node.js runtime to compile
set(NODE_LIB_TYPE "node" CACHE STRING "Determines Node.js runtime's type(node/nw/electron)")
set(NODE_LIB_VERSION "" CACHE STRING "Version value of Node.js runtime library")
set(NODE_LIB_ARCH "x64" CACHE STRING "Node.js runtime architecture")

set(SETUP_NODE_ENVIRONMENT ON CACHE BOOL "Download Node.js runtime and link library during configuration")

message("Current Node.js runtime type: ${NODE_LIB_TYPE} ")
message("Current Node.js runtime version: ${NODE_LIB_VERSION} ")
message("Current Node.js runtime architecture: ${NODE_LIB_ARCH} ")


set(MAIN_SRC
  package.json
  index.js
  stdafx.h 
  stdafx.cpp 
  targetver.h 
  main.cpp
  dllmain.cpp 
)
source_group(main FILES ${MAIN_SRC})

set(WIA_WRAPPER_SRC
  WIADeviceMgr.h 
  WIADeviceMgr.cpp 
)
source_group(wia_wrapper FILES ${WIA_WRAPPER_SRC})

set(UTIL_SRC
  asyncEvent.h 
  asyncEvent.cpp 
  utils.h 
  utils.cpp 
)
source_group(utils FILES ${UTIL_SRC})

# defines library
add_library(wia-scanner-js SHARED
  ${MAIN_SRC}
  ${WIA_WRAPPER_SRC}
  ${UTIL_SRC}
)

# defines precompiled header
add_precompiled_header(wia-scanner-js stdafx.h stdafx.cpp)

set_target_properties(wia-scanner-js PROPERTIES SUFFIX ".node")

#
# Download Node.js runtime in the configuration procedure
#
if(SETUP_NODE_ENVIRONMENT)
  set(CMD_PREFIX "cmd" "/c")

  message(STATUS "Current PATH environment variable is $ENV{PATH}")

  message("Downloading required node-modules....")
  message(STATUS "Working directory is ${CMAKE_CURRENT_SOURCE_DIR}")
  execute_process(COMMAND ${CMD_PREFIX} "npm install" WORKING_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}")

  message("Downloading specified Node.js runtime....")
  if(NOT (${NODE_LIB_TYPE} STREQUAL "node" OR ${NODE_LIB_TYPE} STREQUAL "nw" OR ${NODE_LIB_TYPE} STREQUAL "electron" ))
    message(FATAL_ERROR "Unknown type of Node.js runtime: ${NODE_LIB_NAME}.")
  endif()
  if("${NODE_LIB_VERSION}" STREQUAL "")
    message(FATAL_ERROR "NODE_LIB_VERSION variable is needed.")
  endif()

  execute_process(
    COMMAND ${CMD_PREFIX} "cmake-js install --runtime ${NODE_LIB_TYPE} --runtime-version ${NODE_LIB_VERSION} --arch ia32"
    COMMAND ${CMD_PREFIX} "cmake-js install --runtime ${NODE_LIB_TYPE} --runtime-version ${NODE_LIB_VERSION} --arch x64"
    COMMAND ${CMD_PREFIX} "cmake-js install --runtime ${NODE_LIB_TYPE} --runtime-version ${NODE_LIB_VERSION} --arch arm"
    WORKING_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}"
  )
endif()


#
# Find user home directory path
#
set(USER_HOME_PATH "$ENV{HOMEDRIVE}$ENV{HOMEPATH}")

set(NODE_INCLUDE_PREFIX "${USER_HOME_PATH}/.cmake-js/${NODE_LIB_TYPE}-${NODE_LIB_ARCH}/v${NODE_LIB_VERSION}")

message(STATUS "Current Node.js include/lib dir is ${NODE_INCLUDE_PREFIX}")

# nan(Native Abstractions for Node.js) include dir
set(NAN_INCLUDE_DIR "${CMAKE_CURRENT_SOURCE_DIR}/node_modules/nan")

# Node.js header files dir
if(${NODE_LIB_TYPE} STREQUAL "electron")
  set(NODE_INCLUDES
    "${NODE_INCLUDE_PREFIX}/src"
    "${NODE_INCLUDE_PREFIX}/deps/v8/include"
    "${NODE_INCLUDE_PREFIX}/deps/zlib"
    "${NODE_INCLUDE_PREFIX}/deps/uv/include"
    "${NODE_INCLUDE_PREFIX}/deps/http_parser"
  )
elseif(${NODE_LIB_TYPE} STREQUAL "node")
  set(NODE_INCLUDES
    "${NODE_INCLUDE_PREFIX}/include/node"
  )
elseif(${NODE_LIB_TYPE} STREQUAL "nw")
  set(NODE_INCLUDES
    "${NODE_INCLUDE_PREFIX}/src"
    "${NODE_INCLUDE_PREFIX}/deps/v8/include"
    "${NODE_INCLUDE_PREFIX}/deps/zlib"
    "${NODE_INCLUDE_PREFIX}/deps/uv/include"
    "${NODE_INCLUDE_PREFIX}/deps/openssl"
    "${NODE_INCLUDE_PREFIX}/deps/http_parser"
    "${NODE_INCLUDE_PREFIX}/deps/cares/include"
    "${NODE_INCLUDE_PREFIX}/deps/gtest/include"
    "${NODE_INCLUDE_PREFIX}/deps/icu-small/source"
    "${NODE_INCLUDE_PREFIX}/deps/nghttp2/lib"
  )
else()
  message(FATAL_ERROR "Unknown type of Node.js runtime: ${NODE_LIB_NAME}.")
endif()

target_include_directories(wia-scanner-js PUBLIC 
  ${CMAKE_CURRENT_SOURCE_DIR}
  ${NAN_INCLUDE_DIR}
  ${NODE_INCLUDES}
)

#
# Compile definitions
#
set(MODULE_COMPILE_DEFINITIONS
  UNICODE
  _UNICODE 
)

target_compile_definitions(wia-scanner-js PUBLIC 
  ${MODULE_COMPILE_DEFINITIONS}
)

#
# Link libraries
#
set(MODULE_LINK_LIBRARIES
  "shlwapi.lib"
  "wiaguid.lib"
)

# add node.lib for link under windows
if(WIN32)
  set(NODE_LIB_PATH ${NODE_INCLUDE_PREFIX})

  if(${NODE_LIB_TYPE} STREQUAL "electron")
    # electron
    if(${NODE_LIB_ARCH} STREQUAL "x64")
      string(CONCAT NODE_LIB_PATH ${NODE_LIB_PATH} "/x64")
    endif()
    string(CONCAT NODE_LIB_PATH ${NODE_LIB_PATH} "/node.lib")

    message(STATUS "The Node.js library used to link locates in ${NODE_LIB_PATH}")
    list(APPEND MODULE_LINK_LIBRARIES "${NODE_LIB_PATH}")
  elseif(${NODE_LIB_TYPE} STREQUAL "node")
    # Node.js
    if(${NODE_LIB_ARCH} STREQUAL "ia32")
      string(CONCAT NODE_LIB_PATH ${NODE_LIB_PATH} "/win-x86")
    elseif(${NODE_LIB_ARCH} STREQUAL "x64")
      string(CONCAT NODE_LIB_PATH ${NODE_LIB_PATH} "/win-x64")
    elseif(${NODE_LIB_ARCH} STREQUAL "x64")
      string(CONCAT NODE_LIB_PATH ${NODE_LIB_PATH} "/win-x86")
    endif()

    string(CONCAT NODE_LIB_PATH ${NODE_LIB_PATH} "/node.lib")

    message(STATUS "The Node.js library used to link locates in ${NODE_LIB_PATH}")
    list(APPEND MODULE_LINK_LIBRARIES "${NODE_LIB_PATH}")
  elseif(${NODE_LIB_TYPE} STREQUAL "nw")
    # NW.js
    if(${NODE_LIB_ARCH} STREQUAL "x64")
      string(CONCAT NODE_LIB_PATH ${NODE_LIB_PATH} "/x64")
    endif()
    string(CONCAT NODE_LIB_PATH ${NODE_LIB_PATH} "/node.lib")

    message(STATUS "The Node.js library used to link locates in ${NODE_LIB_PATH}/node.lib")
    list(APPEND MODULE_LINK_LIBRARIES "${NODE_LIB_PATH}/node.lib")
    list(APPEND MODULE_LINK_LIBRARIES "${NODE_LIB_PATH}/nw.lib")
  else()
    message(FATAL_ERROR "Unknown type of Node.js runtime: ${NODE_LIB_NAME}.")
  endif()

endif()

target_link_libraries(wia-scanner-js
  ${MODULE_LINK_LIBRARIES}
)

