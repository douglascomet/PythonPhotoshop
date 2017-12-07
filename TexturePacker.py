#==============================================================================
#!/usr/bin/env python
#title           :TexturePacker.py
#description     :Script used to pack textures using Photoshop
#author          :Doug Halley
#date            :20171204
#version         :3.0
#usage           :Standalone Python Application Executed by TexturePacker.exe
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
    '''
    The class that contains, defines, and creates the UI
    '''

    def __init__(self, parent=None):
        '''
        Class init
        '''
        super(Main, self).__init__(parent)
        self.initUI()

    def initUI(self):
        '''
        Creates UI
        '''

        #==============================================================================
        # Global Variables
        #==============================================================================

        # window title
        self.setWindowTitle('TexturePacker')

        #==============================================================================
        # PYQT Widget Defintions
        #==============================================================================

        # main widget
        self.central_widget = QtGui.QWidget()
        self.central_widget.setLayout(QtGui.QVBoxLayout())

        # Search Directory widget
        self.inputDir_widget = QtGui.QWidget()
        self.inputDir_widget.setLayout(QtGui.QVBoxLayout())

        self.addDirectory = QtGui.QPushButton('Choose Directory')

        self.dirName_lbl = QtGui.QLabel('Selected Directory:')
        self.dirName_lbl.setAlignment(QtCore.Qt.AlignCenter)

        self.dir_lbl = QtGui.QLabel('')
        self.dir_lbl.setAlignment(QtCore.Qt.AlignCenter)

        # Input Channel widget
        self.inputChannel_widget = QtGui.QWidget()
        self.inputChannel_widget.setLayout(QtGui.QVBoxLayout())

        channelInputs_lbl = QtGui.QLabel('Channel Inputs')
        
        # RGBA channel widgets have the same declarations
        self.rChannel_widget = QtGui.QWidget()
        self.rChannel_widget.setLayout(QtGui.QHBoxLayout())

        self.gChannel_widget = QtGui.QWidget()
        self.gChannel_widget.setLayout(QtGui.QHBoxLayout())

        self.bChannel_widget = QtGui.QWidget()
        self.bChannel_widget.setLayout(QtGui.QHBoxLayout())

        self.aChannel_widget = QtGui.QWidget()
        self.aChannel_widget.setLayout(QtGui.QHBoxLayout())

        self.rChannel_lbl = QtGui.QLabel('R Channel')
        self.gChannel_lbl = QtGui.QLabel('G Channel')
        self.bChannel_lbl = QtGui.QLabel('B Channel')

        # QCheckbox is initialized at false
        self.aChannel_checkBox = QtGui.QCheckBox('A Channel')
        self.aChannel_checkBox.setChecked(False)

        self.rChannel_le = QtGui.QLineEdit('')
        self.gChannel_le = QtGui.QLineEdit('')
        self.bChannel_le = QtGui.QLineEdit('')

        # QLineEdit is initialized at false
        self.aChannel_le = QtGui.QLineEdit('')
        self.aChannel_le.setEnabled(False)

        # variable to be used for placeholder text in line edits
        channelSearchText = 'Enter Suffix Here...'

        self.rChannel_le.setPlaceholderText(channelSearchText)
        self.gChannel_le.setPlaceholderText(channelSearchText)
        self.bChannel_le.setPlaceholderText(channelSearchText)
        self.aChannel_le.setPlaceholderText(channelSearchText)

        # creates combobox for year
        self.textureSize_comboBox = QtGui.QComboBox()

        for x in [4096, 2048, 1024, 512, 256, 128, 64]:
            self.textureSize_comboBox.addItem(QtCore.QString(str(x)))

        # packed texture widget, uses similar declaration as RGBA widget declarations
        self.packedTexture_widget = QtGui.QWidget()
        self.packedTexture_widget.setLayout(QtGui.QHBoxLayout())

        self.packedTexture_lbl = QtGui.QLabel('Packed Texture')
        self.packedTexture_le = QtGui.QLineEdit('')
        self.packedTexture_le.setPlaceholderText(channelSearchText)

        # output format widget declaration to contain information string stating the output format
        self.outputFormat_widget = QtGui.QWidget()
        self.outputFormat_widget.setLayout(QtGui.QHBoxLayout())

        self.outputFormatDescription_lbl = QtGui.QLabel('Output Format')
        self.outputFormat_lbl = QtGui.QLabel('24bit .tga')

        self.preProcessTextures_btn = QtGui.QPushButton('Pre-Process Textures')
        self.processTextures_btn = QtGui.QPushButton('Pack Textures')

        #==============================================================================
        # PYQT Widget Assignments
        #==============================================================================

        self.inputDir_widget.layout().addWidget(self.addDirectory)
        self.inputDir_widget.layout().layout().addWidget(self.dirName_lbl)
        self.inputDir_widget.layout().layout().addWidget(self.dir_lbl)

        self.packedTexture_widget.layout().addWidget(self.packedTexture_lbl)
        self.packedTexture_widget.layout().addWidget(self.packedTexture_le)

        self.rChannel_widget.layout().layout().addWidget(self.rChannel_lbl)
        self.rChannel_widget.layout().layout().addWidget(self.rChannel_le)

        self.gChannel_widget.layout().layout().addWidget(self.gChannel_lbl)
        self.gChannel_widget.layout().layout().addWidget(self.gChannel_le)

        self.bChannel_widget.layout().layout().addWidget(self.bChannel_lbl)
        self.bChannel_widget.layout().layout().addWidget(self.bChannel_le)

        self.aChannel_widget.layout().layout().addWidget(self.aChannel_checkBox)
        self.aChannel_widget.layout().layout().addWidget(self.aChannel_le)

        self.inputChannel_widget.layout().addWidget(self.rChannel_widget)
        self.inputChannel_widget.layout().addWidget(self.gChannel_widget)
        self.inputChannel_widget.layout().addWidget(self.bChannel_widget)
        self.inputChannel_widget.layout().addWidget(self.aChannel_widget)

        self.outputFormat_widget.layout().addWidget(self.outputFormatDescription_lbl)
        self.outputFormat_widget.layout().addWidget(self.outputFormat_lbl)

        # adds project widget and tools widget to central widget
        self.central_widget.layout().addWidget(self.inputDir_widget)
        self.central_widget.layout().addWidget(self.inputChannel_widget)
        self.central_widget.layout().addWidget(self.packedTexture_widget)
        self.central_widget.layout().addWidget(self.outputFormat_widget)
        self.central_widget.layout().addWidget(self.processTextures_btn)

        # sets central widget for PyQt window
        self.setCentralWidget(self.central_widget)

        #==============================================================================
        # PYQT Execution Connections
        #==============================================================================

        # triggers for buttons
        self.addDirectory.clicked.connect(lambda: self.getDirectory())
        self.preProcessTextures_btn.clicked.connect(
            lambda: self.preProcessTextures(str(self.dir_lbl.text())))
        self.processTextures_btn.clicked.connect(
            lambda: self.packTextures(str(self.dir_lbl.text())))
        self.aChannel_checkBox.toggled.connect(
            lambda: self.toggleAlphaInput(self.aChannel_checkBox))

    def getDirectory(self):
        '''creates QFileDialog to find designated folder'''
        dlg = QtGui.QFileDialog.getExistingDirectory(
            None, 'Select a folder:', 'C:\\Users\\desktop', QtGui.QFileDialog.ShowDirsOnly)

        self.dir_lbl.setText(dlg)
        print self.dir_lbl.text()

    def toggleAlphaInput(self, checkBox):
        '''toggles alpha channel QLineEdit usablity'''
        if checkBox.isChecked():
            self.aChannel_le.setEnabled(True)
            self.outputFormat_lbl.setText('32bit .tga')
        else:
            self.aChannel_le.setEnabled(False)
            self.outputFormat_lbl.setText('24bit .tga')

    def is_power2(self, num):
        '''# Author: A.Polino - https://code.activestate.com/recipes/577514-chek-if-a-number-is-a-power-of-two/'''

        # states if a number is a power of two
        return num != 0 and ((num & (num - 1)) == 0)

    def preProcessTextures(self, path):
        '''Analyzes textures in folder to determine if their extension is correct'''
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
                    'There were not any files that needed their extensions formatted')
            else:
                self.popupOkWindow(
                    str(count) + 'Files had their extenstions changed to .tga')
            # for x, y in list_of_files.iteritems():
            #     print 'File Name: ', x
            #     print 'File Path: ', y

        else:
            self.popupOkWindow('Invalid Path')

    def packTextures(self, psPath):

        ext = ('.tga', '.png', 'jpg')

        # counterer for number of files
        countTileable = 0

        # variables used to store paths of textures that match desired suffixes
        # variables are initilized to empty strings
        rTexture = ''
        gTexture = ''
        bTexture = ''
        aTexture = ''

        # check to prevent execution with entred output suffix
        if self.packedTexture_le:
            
            # precautionary check to make sure that there is an entry in the RGB channels
            if self.rChannel_le.text() and self.gChannel_le.text() and self.bChannel_le.text():

                # walk through directory, sub directories, and files
                for (dirname, dirs, files) in os.walk(psPath):

                    print dirname

                    if dirs:
                        # print 'dirs'
                        for d in dirs:
                            print d

                            # iterate over entries in a directory
                            for x in os.listdir(os.path.join(dirname, d)):
                                
                                # check if entry in directory is a file
                                if os.path.isfile(os.path.join(os.path.join(dirname, d, x))):
                                    
                                    # check if file extension exists in extension list
                                    if x.lower().endswith(ext):
                                        
                                        # store current path
                                        currentPath = os.path.join(
                                            os.path.join(dirname, d, x))

                                        if str(self.rChannel_le.text()) in x:
                                            print x
                                            rTexture = currentPath

                                        if str(self.bChannel_le.text()) in x:
                                            print x
                                            bTexture = currentPath

                                        if str(self.gChannel_le.text()) in x:
                                            print x
                                            gTexture = currentPath

                                        # similarly to RGB checks, will only check if A QLineEdit is enabled
                                        if self.aChannel_le.isEnabled() and self.aChannel_le.text():
                                            if str(self.aChannel_le.text()) in x:
                                                print x
                                                aTexture = currentPath

                                        # check if a value for R, G, and B is found to continue
                                        if rTexture and gTexture and bTexture:
                                            
                                            # get version of Windows
                                            osVersion = self.checkWindowsVersion()

                                            # open Photoshop
                                            psApp = self.launchPhotoshop(osVersion)

                                            # open texture matching designated suffix to be used for R Channel
                                            r = psApp.Open(rTexture)

                                            # get width and height of texture
                                            docWidth = r.width
                                            docHeight = r.height

                                            # selec and  copy contents of the layer in focus
                                            r.selection.selectAll()
                                            r.activeLayer.Copy()

                                            # use height and width variables to create new texture with same dimentions
                                            blankDoc = psApp.Documents.Add(
                                                docWidth, docHeight, 72, 'new_document', 2, 1, 1)

                                            # blankDoc.channels['Red'] - equivalent to calling channel by name
                                            # activeChannels must receive an array
                                            blankDoc.activeChannels = [blankDoc.channels['Red']]
                                            blankDoc.Paste()

                                            # follows same flow as what was done for R Channel
                                            g = psApp.Open(gTexture)
                                            g.selection.selectAll()
                                            g.activeLayer.Copy()

                                            psApp.activeDocument = blankDoc
                                            blankDoc.activeChannels = [blankDoc.channels['Green']]
                                            blankDoc.Paste()

                                            # follows same flow as what was done for R and G Channels
                                            b = psApp.Open(bTexture)
                                            b.selection.selectAll()
                                            b.activeLayer.Copy()

                                            psApp.activeDocument = blankDoc
                                            blankDoc.activeChannels = [blankDoc.channels['Blue']]
                                            blankDoc.Paste()

                                            # close original textures without saving
                                            r.Close(2)
                                            g.Close(2)
                                            b.Close(2)

                                            # based on earlier A Channel checks, should only proceed if A Channel was desired
                                            if aTexture:
                                                
                                                # follows same flow as what was done for R, G and B Channels
                                                a = psApp.Open(aTexture)
                                                a.selection.selectAll()
                                                a.activeLayer.Copy()

                                                psApp.activeDocument = blankDoc
                                                blankDoc.channels.add()
                                                # blankDoc.Name = 'Alpha 1'
                                                # blankDoc.Kind = 2  # = PsChannelType.psMaskedAreaAlphaChannel
                                                blankDoc.Paste()

                                                a.Close(2)                                            

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
                                            if aTexture:
                                                self.saveTGA(osVersion, psApp, NewFileNamePath, True)
                                            else:
                                                self.saveTGA(osVersion, psApp, NewFileNamePath)

                                            blankDoc.Close(2)

                                            # reset variables after packed texture is exported
                                            rTexture = ''
                                            gTexture = ''
                                            bTexture = ''
                                            aTexture = ''

                            # string variables are reset after traversing all files in a directory
                            rTexture = ''
                            gTexture = ''
                            bTexture = ''
                            aTexture = ''
            else:
                self.popupOkWindow('No Suffix Entred for all RGB Channels')
        else:
            self.popupOkWindow('No Suffix for Packed Texture')
                    
                    
    def checkWindowsVersion(self):
        '''Uses the platform package to determine the version of Windows'''
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
        '''Used to determine if Photoshop is installed and which version is being used'''

        #default Photoshop install path
        if self.osPath('C:\\Program Files\\Adobe\\'):

            # get list of folders in default Photoshop install path
            listdir = self.getPathDirs('C:\\Program Files\\Adobe\\')

            # Searches Adobe directory to determine to see if a version of Photoshop is installed
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

        if osVer == '10':
            psApp = comtypes.client.CreateObject(
                'Photoshop.Application', dynamic=True)
            psApp.Visible = True

            # Set the default unit to pixels!
            psApp.Preferences.RulerUnits = 1

            return psApp
        else:
            self.popupOkWindow('Error with determining OS Version to launch Photoshop')

    def saveTGA(self, osVer, doc, tgaFile, saveAlpha=False):

        # check if currently on Win10
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
        ''' Generic popup window with an OK button and can display message based on use '''

        popupWindow = QtGui.QMessageBox()

        popupWindow.setText('Textures Found')
        popupWindow.setDetailedText(str(message))
        popupWindow.setStandardButtons(QtGui.QMessageBox.Ok)

        popupWindow.exec_()

    def popupOkWindow(self, message):
        ''' Generic popup window with an OK button and can display message based on use '''

        popupWindow = QtGui.QMessageBox()
        popupWindow.setText(message)
        popupWindow.setStandardButtons(QtGui.QMessageBox.Ok)

        popupWindow.exec_()

    def getPathDirs(self, filePath):
        '''gets list of directories if path exists'''
        if self.osPath(filePath):
            return os.listdir(filePath)

    
    def osPath(self, filePath):
        '''check if path exists'''
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
