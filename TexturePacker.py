#==============================================================================
#!/usr/bin/env python
#title           :TexturePacker.py
#description     :Script used to pack textures using Photoshop
#author          :Doug Halley
#date            :20171204
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
        self.setWindowTitle("TexturePacker")

        #==============================================================================
        # PYQT Widget Defintions
        #==============================================================================

        # main widget
        self.centralWidget = QtGui.QWidget()
        self.centralWidget.setLayout(QtGui.QVBoxLayout())

        self.addDirectory = QtGui.QPushButton("Choose Directory")

        self.inputDir_widget = QtGui.QWidget()
        self.inputDir_widget.setLayout(QtGui.QVBoxLayout())

        self.dirName_lbl = QtGui.QLabel("Selected Directory:")
        self.dirName_lbl.setAlignment(QtCore.Qt.AlignCenter)

        self.dir_lbl = QtGui.QLabel("")
        self.dir_lbl.setAlignment(QtCore.Qt.AlignCenter)

        self.inputChannel_widget = QtGui.QWidget()
        self.inputChannel_widget.setLayout(QtGui.QVBoxLayout())

        channelInputs_lbl = QtGui.QLabel("Channel Inputs")

        self.rChannel_widget = QtGui.QWidget()
        self.rChannel_widget.setLayout(QtGui.QHBoxLayout())

        self.gChannel_widget = QtGui.QWidget()
        self.gChannel_widget.setLayout(QtGui.QHBoxLayout())

        self.bChannel_widget = QtGui.QWidget()
        self.bChannel_widget.setLayout(QtGui.QHBoxLayout())

        self.aChannel_widget = QtGui.QWidget()
        self.aChannel_widget.setLayout(QtGui.QHBoxLayout())

        self.rChannel_lbl = QtGui.QLabel("R Channel")
        self.gChannel_lbl = QtGui.QLabel("G Channel")
        self.bChannel_lbl = QtGui.QLabel("B Channel")
        self.aChannel_checkBox = QtGui.QCheckBox("A Channel")
        self.aChannel_checkBox.setChecked(False)

        self.rChannel_le = QtGui.QLineEdit("")
        self.gChannel_le = QtGui.QLineEdit("")
        self.bChannel_le = QtGui.QLineEdit("")
        self.aChannel_le = QtGui.QLineEdit("")
        self.aChannel_le.setEnabled(False)

        channelSearchText = 'Enter Prefix or Suffix'

        self.rChannel_le.setPlaceholderText(channelSearchText)
        self.gChannel_le.setPlaceholderText(channelSearchText)
        self.bChannel_le.setPlaceholderText(channelSearchText)
        self.aChannel_le.setPlaceholderText(channelSearchText)

        self.rChannel_widget.layout().layout().addWidget(self.rChannel_lbl)
        self.rChannel_widget.layout().layout().addWidget(self.rChannel_le)

        self.gChannel_widget.layout().layout().addWidget(self.gChannel_lbl)
        self.gChannel_widget.layout().layout().addWidget(self.gChannel_le)

        self.bChannel_widget.layout().layout().addWidget(self.bChannel_lbl)
        self.bChannel_widget.layout().layout().addWidget(self.bChannel_le)

        self.aChannel_widget.layout().layout().addWidget(self.aChannel_checkBox)
        self.aChannel_widget.layout().layout().addWidget(self.aChannel_le)

        # creates combobox for year
        self.textureSize_comboBox = QtGui.QComboBox()

        for x in [4096, 2048, 1024, 512, 256, 128, 64]:
            self.textureSize_comboBox.addItem(QtCore.QString(str(x)))

        self.packedTexture_widget = QtGui.QWidget()
        self.packedTexture_widget.setLayout(QtGui.QHBoxLayout())

        self.packedTexture_lbl = QtGui.QLabel("Packed Texture")
        self.packedTexture_le = QtGui.QLineEdit("")
        self.packedTexture_le.setPlaceholderText('Enter Suffix Here')

        self.packedTexture_widget.layout().addWidget(self.packedTexture_lbl)
        self.packedTexture_widget.layout().addWidget(self.packedTexture_le)

        self.preProcessTextures_btn = QtGui.QPushButton("Pre-Process Textures")
        self.processTextures_btn = QtGui.QPushButton("Process Textures")

        self.inputDir_widget.layout().layout().addWidget(self.dirName_lbl)
        self.inputDir_widget.layout().layout().addWidget(self.dir_lbl)

        self.inputChannel_widget.layout().addWidget(self.rChannel_widget)
        self.inputChannel_widget.layout().addWidget(self.gChannel_widget)
        self.inputChannel_widget.layout().addWidget(self.bChannel_widget)
        self.inputChannel_widget.layout().addWidget(self.aChannel_widget)
        # self.inputChannel_widget.layout().layout().addWidget(self.textureSize_comboBox)

        # adds project widget and tools widget to central widget
        self.centralWidget.layout().addWidget(self.addDirectory)
        self.centralWidget.layout().addWidget(self.inputDir_widget)
        self.centralWidget.layout().addWidget(self.inputChannel_widget)
        self.centralWidget.layout().addWidget(self.packedTexture_widget)
        self.centralWidget.layout().addWidget(self.processTextures_btn)

        # sets central widget for PyQt window
        self.setCentralWidget(self.centralWidget)

        #==============================================================================
        # PYQT Execution Connections
        #==============================================================================

        # triggers for buttons
        self.addDirectory.clicked.connect(lambda: self.getDirectory())
        self.preProcessTextures_btn.clicked.connect(
            lambda: self.preProcessTextures(str(self.dir_lbl.text())))
        self.processTextures_btn.clicked.connect(
            lambda: self.textureResize(str(self.dir_lbl.text())))
        self.aChannel_checkBox.toggled.connect(
            lambda: self.toggleAlphaInput(self.aChannel_checkBox))

    # creates QFileDialog to find designated folder
    def getDirectory(self):
        dlg = QtGui.QFileDialog.getExistingDirectory(
            None, 'Select a folder:', 'C:\\Users\\desktop', QtGui.QFileDialog.ShowDirsOnly)

        self.dir_lbl.setText(dlg)
        print self.dir_lbl.text()

    def toggleAlphaInput(self, checkBox):

        if checkBox.isChecked():
            self.aChannel_le.setEnabled(True)
        else:
            self.aChannel_le.setEnabled(False)

    # Author: A.Polino - https://code.activestate.com/recipes/577514-chek-if-a-number-is-a-power-of-two/
    def is_power2(self, num):

        # states if a number is a power of two
        return num != 0 and ((num & (num - 1)) == 0)

    def preProcessTextures(self, path):
        #import collections
        #path = "C:\\Users\\dhalley\\Desktop\\GarageScene"
        #list_of_files = collections.OrderedDict()
        count = 0

        if self.osPath(path):

            for (dirpath, dirnames, filenames) in os.walk(path):
                for filename in filenames:
                    if filename.endswith('.targa') or filename.endswith('.TGA'):

                        print filename
                        base_file, ext = os.path.splitext(filename)
                        print base_file
                        print type(base_file)
                        OriginalFileNamePath = os.path.join(dirpath, filename)
                        print OriginalFileNamePath
                        NewFileNamePath = os.path.join(
                            dirpath, base_file + '.tga')
                        #in order to successfully rename a file you need the file joined with the path
                        os.rename(OriginalFileNamePath, NewFileNamePath)
                        list_of_files[filename] = os.sep.join(
                            [dirpath, filename])
                        count += 1
            if count == 0:
                self.popupOkWindow(
                    "There were not any files that needed their extensions formatted")
            else:
                self.popupOkWindow(
                    str(count) + "Files had their extenstions changed to .tga")
            # for x, y in list_of_files.iteritems():
            #     print 'File Name: ', x
            #     print 'File Path: ', y

        else:
            self.popupOkWindow('Invalid Path')

    def textureResize(self, path):

        psFileLocation = path

        ext = (".tga", ".png")

        # counterer for number of files
        countTileable = 0

        rFound = ''
        gFound = ''
        bFound = ''
        aFound = ''

        # walk through directory, sub directories, and files
        for (dirname, dirs, files) in os.walk(psFileLocation):

            print dirname
            # if not dirs:
            #     print 'asdfasdfasdf' + str(dirs)
            # print files
            if dirs:
                # print 'dirs'
                for d in dirs:
                    print d
                    for x in os.listdir(os.path.join(dirname, d)):
                        if os.path.isfile( os.path.join( os.path.join( dirname, d, x ) ) ):
                            # check if file extension exists in extension list
                            if x.lower().endswith('.tga'):
                                
                                #store current path
                                currentPath = os.path.join(
                                    os.path.join(dirname, d, x))

                                if self.rChannel_le.text():
                                    if str(self.rChannel_le.text()) in x:
                                        print x
                                        rFound = currentPath

                                if self.bChannel_le.text():
                                    if str(self.bChannel_le.text()) in x:
                                        print x
                                        bFound = currentPath

                                if self.gChannel_le.text():
                                    if str(self.gChannel_le.text()) in x:
                                        print x
                                        gFound = currentPath

                                if self.aChannel_le.text():
                                    if str(self.aChannel_le.text()) in x:
                                        print x
                                        aFound = currentPath

                                # check if a value for R, G, and B is found
                                if rFound and gFound and bFound:
                                    
                                    # get used version of Windows
                                    osVersion = self.checkWindowsVersion()

                                    # open Photoshop
                                    psApp = self.launchPhotoshop(osVersion)

                                    r = psApp.Open(rFound)
                                    docWidth = r.width
                                    docHeight = r.height
                                    r.selection.selectAll()
                                    r.activeLayer.Copy()

                                    blankDoc = psApp.Documents.Add(
                                        docWidth, docHeight, 72, "new_document", 2, 1, 1)

                                    # blankDoc.channels['Red'] - equivalent to calling channel by name
                                    # activeChannels must receive an array
                                    blankDoc.activeChannels = [blankDoc.channels['Red']]
                                    blankDoc.Paste()

                                    g = psApp.Open(gFound)
                                    g.selection.selectAll()
                                    g.activeLayer.Copy()

                                    psApp.activeDocument = blankDoc
                                    blankDoc.activeChannels = [blankDoc.channels['Green']]
                                    blankDoc.Paste()

                                    b = psApp.Open(bFound)
                                    b.selection.selectAll()
                                    b.activeLayer.Copy()

                                    psApp.activeDocument = blankDoc
                                    blankDoc.activeChannels = [blankDoc.channels['Blue']]
                                    blankDoc.Paste()

                                    # close original version without saving
                                    r.Close(2)
                                    g.Close(2)
                                    b.Close(2)

                                    if aFound:
                                        a = psApp.Open(aFound)
                                        a.selection.selectAll()
                                        a.activeLayer.Copy()

                                        psApp.activeDocument = blankDoc
                                        blankDoc.channels.add()
                                        # blankDoc.Name = 'Alpha 1'
                                        # blankDoc.Kind = 2  # = PsChannelType.psMaskedAreaAlphaChannel
                                        blankDoc.Paste()

                                        a.Close(2)

                                    if self.packedTexture_le:

                                        splitPath, splitPathFileName = os.path.split(currentPath)

                                        fileName, fileExt = os.path.splitext(splitPathFileName)
                                        
                                        # gets first element of split from '_' based on naming convention at BBA
                                        splitFileName = fileName.split('_')

                                        if len(splitFileName) > 1:
                                            splitFileName.pop()
                                            splitFileName = '_'.join(splitFileName)
                                        else:
                                            splitFileName = splitFileName[0]

                                        newFileName = str(splitFileName) + str(self.packedTexture_le.text()) + fileExt

                                        NewFileNamePath = os.path.join(splitPath, newFileName)

                                        # if there is an alpha input be sure to export TGA with alpha option on
                                        if aFound:
                                            self.saveTGA(osVersion, psApp, NewFileNamePath, True)
                                        else:
                                            self.saveTGA(osVersion, psApp, NewFileNamePath)

                                        blankDoc.Close(2)

                                        # reset variables after packed texture is exported
                                        rFound = ''
                                        gFound = ''
                                        bFound = ''
                                        aFound = ''
                                    else:
                                        self.popupOkWindow('No Suffix for Packed Texture')
                        
                    print '^ Skipped'
                    rFound = ''
                    gFound = ''
                    bFound = ''
                    aFound = ''
                    
    def checkWindowsVersion(self):
        import platform
        currentPlatform = platform.platform()

        if 'Windows' in currentPlatform:
            #splits windows version based on dashes
            splitPlatformName = currentPlatform.split('-')
            #returns windows version number
            return splitPlatformName[1]
        else:
            self.popupOkWindow('Untested OS. Tool only works on Windows')

    def checkPhotoshopVersion(self):

        #default Photoshop install path
        if self.osPath('C:\\Program Files\\Adobe\\'):

            # get list of folders in default Photoshop install path
            listdir = self.getPath('C:\\Program Files\\Adobe\\')

            # determine to see if a version of Photoshop is installed
            foundItems = [x for x in listdir if 'Photoshop' in x]

            if foundItems:

                # check how many versions of Photoshop are installed
                if len(foundItems) == 1:
                    foundPhotoshop = str(foundItems[0])

                    # check if a CC version of Photoshop is installed
                    if 'CC' in foundPhotoshop:

                        # remove spaces from version of Photoshop
                        splitPhotoshop = foundPhotoshop.split(' ')

                        # get last element from split list
                        verNumber = splitPhotoshop[-1]

                        # create empty string to store value and testing
                        version = ''

                        # version 14
                        if verNumber == 'CC':
                            version = '70'
                        elif int(verNumber):
                            # version 15
                            if verNumber == '2014':
                                version = '80'
                            # version 16
                            elif verNumber == '2015':
                                version = '90'
                            # version 17
                            elif verNumber == '2016':
                                version = '100'
                            # version 18
                            elif verNumber == '2017':
                                version = '110'
                            # version 19
                            elif verNumber == '2018':
                                version = '120'

                        if version:
                            return version
                        else:
                            self.popupOkWindow(
                                'Error getting installed Photoshop Version')
                    else:
                        # if version.startswith('12'):
                        #     version = '12'
                        # elif version.startswith('13'):
                        #     version = '13'
                        self.popupOkWindow('Non CC Version of PS')
                else:
                    self.popupOkWindow(
                        'Multiple Versions of Photoshop installed')
            else:
                self.popupOkWindow('Photoshop not installed')
        else:
            self.popupOkWindow(
                'Adobe Software not installed in default directory')

    def launchPhotoshop(self, osVer):

        # if osVer == '10':
        psApp = comtypes.client.CreateObject(
            'Photoshop.Application', dynamic=True)
        psApp.Visible = True

        #Set the default unit to pixels!
        psApp.Preferences.RulerUnits = 1

        return psApp

    def saveTGA(self, osVer, doc, tgaFile, saveAlpha=False):

        if osVer == '10':
            tgaOptions = comtypes.client.CreateObject(
                'Photoshop.TargaSaveOptions', dynamic=True)
            tgaOptions.Resolution = 24
            tgaOptions.AlphaChannels = False
            tgaOptions.RLECompression = False

            if saveAlpha:
                tgaOptions.Resolution = 32
                tgaOptions.AlphaChannels = True

            doc.ActiveDocument.SaveAs(tgaFile, tgaOptions, True)

    def popupDetailedOkWindow(self, message):
        """ Generic popup window with an OK button and can display message based on use """

        popupWindow = QtGui.QMessageBox()

        popupWindow.setText('Textures Found')
        popupWindow.setDetailedText(str(message))
        popupWindow.setStandardButtons(QtGui.QMessageBox.Ok)

        popupWindow.exec_()

    def popupOkWindow(self, message):
        """ Generic popup window with an OK button and can display message based on use """

        popupWindow = QtGui.QMessageBox()
        popupWindow.setText(message)
        popupWindow.setStandardButtons(QtGui.QMessageBox.Ok)

        popupWindow.exec_()

    #get list of directories if path exists
    def getPath(self, filePath):
        if self.osPath(filePath):
            return os.listdir(filePath)

    #determine if path exists
    def osPath(self, filePath):
        #print filePath
        #print type(filePath)
        if os.path.isdir(filePath):
            return True
        else:
            return False


if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)
    myWidget = Main()
    myWidget.show()
    sys.exit(app.exec_())
