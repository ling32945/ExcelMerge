#-------------------------------------------------
#
# Project created by QtCreator 2018-03-22T19:28:39
#
#-------------------------------------------------

QT       += core gui

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = ExcelMerge
TEMPLATE = app

# The following define makes your compiler emit warnings if you use
# any feature of Qt which has been marked as deprecated (the exact warnings
# depend on your compiler). Please consult the documentation of the
# deprecated API in order to know how to port your code away from it.
DEFINES += QT_DEPRECATED_WARNINGS

# You can also make your code fail to compile if you use deprecated APIs.
# In order to do so, uncomment the following line.
# You can also select to disable deprecated APIs only up to a certain version of Qt.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0


SOURCES += \
        main.cpp \
        mainwindow.cpp

HEADERS += \
        mainwindow.h

FORMS += \
        mainwindow.ui


LIBXL_LIB_PATH = $$PWD/../../Libxl
INCLUDEPATH += $$LIBXL_LIB_PATH/include_cpp

macx {
    LIBS += -framework LibXL

    #DEPENDPATH += $$LIBXL_LIB_PATH
    QMAKE_LFLAGS += -F$$LIBXL_LIB_PATH/
    QMAKE_POST_LINK +=$$quote(mkdir $${TARGET}.app/Contents/Frameworks;cp -R $${LIBXL_LIB_PATH}/LibXL.framework $${TARGET}.app/Contents/Frameworks/)
}
else{
    LIBS += -L$$LIBXL_LIB_PATH/lib64/ -llibxl
    #DEPENDPATH += $$LIBXL_LIB_PATH/bin64
    #QMAKE_POST_LINK += $$quote(cmd echo $${PWD})
    QMAKE_POST_LINK +=$$quote(cmd /c copy /y ..\..\Libxl\bin64\libxl.dll .\debug)
}

RESOURCES += \
    image.qrc


