#==============================================================================
#!/usr/bin/env python
#title           :TextureResizer.py
#description     :Script used to resize images within a directory by using Photoshop
#author          :Doug Halley
#date            :20171112
#version         :1.0
#usage           :Standalone Python Application Executed by TextureResizer.exe
#notes           :
#python_version  :2.7.14
#pyqt_version    :4.11.4
#==============================================================================

import os
import sys
import comtypes.client

from _ctypes import COMError
from PIL import Image

from PyQt4 import QtGui
from PyQt4 import QtCore

class Main(QtGui.QMainWindow):
    """
    The class that contains, defines, and creates the UI
    """

    def __init__(self, parent=None):
        """
        Class init
        """
        super(Main, self).__init__(parent)
        self.initUI()

    def initUI(self):
        """
        Creates UI
        """

        #==============================================================================
        # Global Variables
        #==============================================================================

        # window title
        self.setWindowTitle("TextureResizer")

        #==============================================================================
        # PYQT Widget Defintions
        #==============================================================================

        # main widget
        self.centralWidget = QtGui.QWidget()
        self.centralWidget.setLayout(QtGui.QVBoxLayout())

        self.addDirectory = QtGui.QPushButton("Choose Directory")

        # widget for project location radio buttons and year combobox
        self.textureSizeWidget = QtGui.QWidget()
        self.textureSizeWidget.setLayout(QtGui.QHBoxLayout())

        textureSize_lbl = QtGui.QLabel("Texture Size")

        # creates combobox for year
        self.textureSize_comboBox = QtGui.QComboBox()

        for x in [4096, 2048, 1024, 512, 256, 128, 64]:
            self.textureSize_comboBox.addItem(QtCore.QString(str(x)))

        self.dirName_lbl = QtGui.QLabel("")
        self.dirName_lbl.setAlignment(QtCore.Qt.AlignCenter)

        self.processTextures = QtGui.QPushButton("Process Textures")

        self.textureSizeWidget.layout().layout().addWidget(textureSize_lbl)
        self.textureSizeWidget.layout().layout().addWidget(self.textureSize_comboBox)

        # adds project widget and tools widget to central widget
        self.centralWidget.layout().addWidget(self.addDirectory)
        self.centralWidget.layout().addWidget(self.dirName_lbl)
        self.centralWidget.layout().addWidget(self.textureSizeWidget)
        self.centralWidget.layout().addWidget(self.processTextures)
        
        # sets central widget for PyQt window
        self.setCentralWidget(self.centralWidget)

        #==============================================================================
        # PYQT Execution Connections
        #==============================================================================

        # triggers for buttons
        self.addDirectory.clicked.connect(lambda: self.getDirectory())
        self.processTextures.clicked.connect(lambda: self.textureResize())

    # creates QFileDialog to find designated folder
    def getDirectory(self):
        dlg = QtGui.QFileDialog.getExistingDirectory(
            None, 'Select a folder:', 'C:\\Users\\desktop', QtGui.QFileDialog.ShowDirsOnly)

        self.dirName_lbl.setText(dlg)


    # Author: A.Polino - https://code.activestate.com/recipes/577514-chek-if-a-number-is-a-power-of-two/
    def is_power2(self, num):

        # states if a number is a power of two
        return num != 0 and ((num & (num - 1)) == 0)

    def textureResize(self):

        psFileLocation = "C:\\Users\dhalley\\Desktop\\GarageScene"

        ext = (".tga", ".png")

        # counterer for number of files
        countTileable = 0

        # variable to check size of images by
        targetSize = 512

        # list used to collect images larger than targetSize
        largerThanTargetSize = []

        testPrint = ''

        # walk through directory, sub directories, and files
        for ( dirname, dirs, files ) in os.walk( psFileLocation ):
            
            #print dirname
            #print dirs
            #print files

            # iterate over list of subdirectories
            for d in dirs:
                
                # precautionary check to ensure is valid path
                if os.path.isdir( os.path.join( dirname, d ) ):

                    # iterate over entries in found subdirectory
                    for x in os.listdir( os.path.join( dirname, d ) ):
                            
                        # check if x during iteration is a file
                        if os.path.isfile( os.path.join( os.path.join( dirname, d, x ) ) ):
                            
                            # check if file extension exists in extension list
                            if x.lower().endswith( ext ):
                                


                                # define imagePath string
                                imagePath = os.path.join( dirname, d, x )

                                # use PIL to create Image object
                                with Image.open(imagePath) as im:
                                    sizeOfImage = im.size

                                # sizeOfImage returns tuple (width, height)
                                # check that image is square by comparing width and height
                                if sizeOfImage[0] == sizeOfImage[ 1 ]:
                                    
                                    # if width/height are equal, use either value to check if power of 2
                                    if self.is_power2(sizeOfImage[ 0 ]):
                                        print x + ' - ' + '{0}'.format( sizeOfImage )
                                        countTileable = countTileable + 1

                                        if sizeOfImage[ 0 ] > targetSize:
                                            testPrint = testPrint + imagePath + '\n'
                                            largerThanTargetSize.append( imagePath )
                                

        psApp = comtypes.client.CreateObject('Photoshop.Application', dynamic=True)
        psApp.Visible = True

        #Set the default unit to pixels!
        psApp.Preferences.RulerUnits = 1

        test = psApp.Open(largerThanTargetSize[0])
        # psApp.Application.ActiveDocument
        test.resizeImage(targetSize, targetSize)

        psApp.activeDocument = test

        tgaOptions = comtypes.client.CreateObject(
            'Photoshop.TargaSaveOptions', dynamic=True)
        tgaOptions.Resolution = 24
        tgaOptions.AlphaChannels = False
        tgaOptions.RLECompression = False

        filename, file_extension = os.path.splitext(largerThanTargetSize[0])

        newFileName = filename + '_' + \
            str(targetSize) + str(2) + file_extension

        test.SaveAs(newFileName, tgaOptions, True)

        # print 'Number of Tileable Files:', countTileable
        # print 'Large Files List Size:', len( largerThanTargetSize )
        # # print 'Large Files List:', largerThanTargetSize

        
        # psApp.Open(largerThanTargetSize[0])
        # test=psApp.Application.ActiveDocument
        # test.resizeImage(targetSize, targetSize)

        # filename, file_extension = os.path.splitext(largerThanTargetSize[0])

        # newFileName = filename + '_' + str(targetSize) + str(2)

        # print newFileName + file_extension
        # psApp.activeDocument = test

        # try:
        #     psApp.activeDocument.SaveAs((newFileName + file_extension), targaSettings, True, 2)
        # except COMError as e:
        #     print e

    def launchPhotoshop(self):
        
        psApp = comtypes.client.CreateObject('Photoshop.Application')
        psApp.Visible = True

        #Set the default unit to pixels!
        psApp.Preferences.RulerUnits = 1

        return psApp

    def saveTGA(self, doc, tgaFile, saveAlpha=False):
        
        tgaOptions = comtypes.client.CreateObject(
            'Photoshop.TargaSaveOptions')
        tgaOptions.Resolution = 24
        tgaOptions.AlphaChannels = False
        tgaOptions.RLECompression = False

        if saveAlpha:
            tgaOptions.Resolution = 32
            tgaOptions.AlphaChannels = True

        doc.SaveAs(tgaFile, tgaOptions, True)

    def popupOkWindow(self, message):
        """ Generic popup window with an OK button and can display message based on use """

        popupWindow = QtGui.QMessageBox()

        popupWindow.setText('Textures Found')
        popupWindow.setDetailedText(str(message))
        popupWindow.setStandardButtons(QtGui.QMessageBox.Ok)

        popupWindow.exec_()


if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)
    myWidget = Main()
    myWidget.show()
    sys.exit(app.exec_())
