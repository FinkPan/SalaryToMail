cmake_minimum_required(VERSION 3.9.0)

project(SalaryToMail)

include_directories("src")

set(CMAKE_AUTOMOC ON)
set(CMAKE_INCLUDE_CURRENT_DIR ON)
set(CMAKE_MODULE_PATH ${CMAKE_MODULE_PATH} "${CMAKE_SOURCE_DIR}/cmake/Modules/")

find_package(Qt5Widgets REQUIRED)
find_package(Qt5AxContainer REQUIRED)
find_package(Qt5Core REQUIRED)
get_target_property(QtCore_location Qt5::Core LOCATION)


qt5_wrap_ui(UI_FILES ui/mainwindow.ui)
qt5_add_resources(RES_FILES "res/res.qrc")

set(SOURCE_FILES
      src/main.cpp
      src/mainwindow.hpp
      src/mainwindow.cpp
      src/readexcelfile.hpp
      src/readexcelfile.cpp
      src/data.hpp
      src/data.cpp
    )

add_executable(SalaryToMail ${SOURCE_FILES} ${UI_FILES} ${RES_FILES})

target_link_libraries(SalaryToMail ${Qt5Widgets_LIBRARIES} ${Qt5AxContainer_LIBRARIES})
