# ==============================================================================
# !/usr/bin/env python
# title           :Pyotoshop.py
# description     :Script used to automate Photoshop operations across several files
# author          :Doug Halley
# date            :20171206
# version         :5.0
# usage           :Standalone Python Application Executed by Pyotoshop.exe
# notes           :
# python_version  :2.7.14
# pyqt_version    :4.11.4
# ==============================================================================

import os
import sys
import comtypes.client

from _ctypes import COMError
from PIL import Image

from PyQt4 import QtGui
from PyQt4 import QtCore

class Main(QtGui.QMainWindow):
    '''
    The class that contains, defines, and creates the UI.
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

        # ==============================================================================
        # Global Variables
        # ==============================================================================

        self.extensions = ('.tga', '.png', '.jpg')
        self.textureSizes = (4096, 2048, 1024, 512, 256, 128, 64)

        # ==============================================================================
        # PYQT Widget Defintions
        # ==============================================================================

        # window title
        self.setWindowTitle('Pyotoshop')

        # main widget ------------------------------------------------------------------
        central_widget = QtGui.QWidget()
        central_widget.setLayout(QtGui.QVBoxLayout())

        # Search Directory widget ------------------------------------------------------
        inputDir_widget = QtGui.QWidget()
        inputDir_widget.setLayout(QtGui.QVBoxLayout())

        self.addDirectory = QtGui.QPushButton('Choose Directory')

        dirName_lbl = QtGui.QLabel('Selected Directory:')
        dirName_lbl.setAlignment(QtCore.Qt.AlignCenter)

        self.dir_lbl = QtGui.QLabel('')
        self.dir_lbl.setAlignment(QtCore.Qt.AlignCenter)

        # widget to fix extensions for targa files --------------------------------------
        texturePreProcess_widget = QtGui.QWidget()
        texturePreProcess_widget.setLayout(QtGui.QVBoxLayout())

        self.preProcessTextures_btn = QtGui.QPushButton('Pre-Process Textures')

        # parent widget for texture resizer elements ------------------------------------
        textureResizer_widget = QtGui.QWidget()
        textureResizer_widget.setLayout(QtGui.QVBoxLayout())

        # widget for combobox that contains varying texture sizes -----------------------
        textureSize_widget = QtGui.QWidget()
        textureSize_widget.setLayout(QtGui.QHBoxLayout())

        textureSize_lbl = QtGui.QLabel('Texture Size')

        # creates combobox for texture sizes
        self.textureSize_comboBox = QtGui.QComboBox()

        for x in self.textureSizes:
            self.textureSize_comboBox.addItem(QtCore.QString(str(x)))

        self.processTextures_btn = QtGui.QPushButton(
                        'Process Textures')

        # Texture Packer widget to contain all elements need to run this tool ------------
        texturePacker_widget = QtGui.QWidget()
        texturePacker_widget.setLayout(QtGui.QVBoxLayout())

        # Input Channel widget, child widget of texturePacker_widget ---------------------
        inputChannel_widget = QtGui.QWidget()
        inputChannel_widget.setLayout(QtGui.QVBoxLayout())

        channelInputs_lbl = QtGui.QLabel('Channel Inputs')

        # RGBA channel widgets have the same declarations --------------------------------
        # child widgets of inputChannel_widget -------------------------------------------
        rChannel_widget = QtGui.QWidget()
        rChannel_widget.setLayout(QtGui.QHBoxLayout())

        gChannel_widget = QtGui.QWidget()
        gChannel_widget.setLayout(QtGui.QHBoxLayout())

        bChannel_widget = QtGui.QWidget()
        bChannel_widget.setLayout(QtGui.QHBoxLayout())

        aChannel_widget = QtGui.QWidget()
        aChannel_widget.setLayout(QtGui.QHBoxLayout())

        rChannel_lbl = QtGui.QLabel('R Channel')
        gChannel_lbl = QtGui.QLabel('G Channel')
        bChannel_lbl = QtGui.QLabel('B Channel')

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

        # packed texture widget, uses similar declaration as RGBA widget declarations ----
        # child widgets of inputChannel_widget -------------------------------------------
        packedTexture_widget = QtGui.QWidget()
        packedTexture_widget.setLayout(QtGui.QHBoxLayout())

        packedTexture_lbl = QtGui.QLabel('Packed Texture')
        self.packedTexture_le = QtGui.QLineEdit('')
        self.packedTexture_le.setPlaceholderText(channelSearchText)

        # output format widget declaration to contain string stating the output format ---
        # child widgets of inputChannel_widget -------------------------------------------
        outputFormat_widget = QtGui.QWidget()
        outputFormat_widget.setLayout(QtGui.QHBoxLayout())

        outputFormatDescription_lbl = QtGui.QLabel('Output Format')
        self.outputFormat_lbl = QtGui.QLabel('24bit .tga')

        self.packTextures_btn = QtGui.QPushButton('Pack Textures')

        # tab widget to contain all the subset tools -------------------------------------
        # child widgets of central_widget ------------------------------------------------
        self.textureTools_tab_widget = QtGui.QTabWidget()
        self.textureTools_tab_widget.setLayout(QtGui.QHBoxLayout())

        self.textureTools_tab_widget.addTab(texturePacker_widget, 'Pack Textures')
        self.textureTools_tab_widget.addTab(textureResizer_widget, 'Resize Textures')
        # self.textureTools_tab_widget.addTab(texturePreProcess_widget, 'Texture PreProcesser')
        self.textureTools_tab_widget.setEnabled(False)

        # ================================================================================
        # PYQT Widget Assignments
        # ================================================================================

        # Assigments for process textures widget -----------------------------------------
        texturePreProcess_widget.layout().layout().addWidget(self.preProcessTextures_btn)

        # Assigments for texture resizer widget ------------------------------------------
        textureSize_widget.layout().layout().addWidget(textureSize_lbl)
        textureSize_widget.layout().layout().addWidget(self.textureSize_comboBox)

        # Assigments for texture packer parent widget ------------------------------------
        textureResizer_widget.layout().layout().addWidget(textureSize_widget)
        textureResizer_widget.layout().addWidget(self.processTextures_btn)

        # Assigments for texture packer childeren widgets --------------------------------
        inputDir_widget.layout().addWidget(self.addDirectory)
        inputDir_widget.layout().layout().addWidget(dirName_lbl)
        inputDir_widget.layout().layout().addWidget(self.dir_lbl)

        packedTexture_widget.layout().addWidget(packedTexture_lbl)
        packedTexture_widget.layout().addWidget(self.packedTexture_le)

        rChannel_widget.layout().layout().addWidget(rChannel_lbl)
        rChannel_widget.layout().layout().addWidget(self.rChannel_le)

        gChannel_widget.layout().layout().addWidget(gChannel_lbl)
        gChannel_widget.layout().layout().addWidget(self.gChannel_le)

        bChannel_widget.layout().layout().addWidget(bChannel_lbl)
        bChannel_widget.layout().layout().addWidget(self.bChannel_le)

        aChannel_widget.layout().layout().addWidget(self.aChannel_checkBox)
        aChannel_widget.layout().layout().addWidget(self.aChannel_le)

        inputChannel_widget.layout().addWidget(rChannel_widget)
        inputChannel_widget.layout().addWidget(gChannel_widget)
        inputChannel_widget.layout().addWidget(bChannel_widget)
        inputChannel_widget.layout().addWidget(aChannel_widget)

        outputFormat_widget.layout().addWidget(outputFormatDescription_lbl)
        outputFormat_widget.layout().addWidget(self.outputFormat_lbl)

        texturePacker_widget.layout().addWidget(inputChannel_widget)
        texturePacker_widget.layout().addWidget(packedTexture_widget)
        texturePacker_widget.layout().addWidget(outputFormat_widget)
        texturePacker_widget.layout().addWidget(self.packTextures_btn)

        # adds directiory widget and tool widgets to central widget ---------------------
        central_widget.layout().addWidget(inputDir_widget)
        central_widget.layout().addWidget(self.textureTools_tab_widget)

        # sets central widget for PyQt window
        self.setCentralWidget(central_widget)

        # ================================================================================
        # PYQT Execution Connections
        # ================================================================================

        # triggers for buttons
        self.addDirectory.clicked.connect(lambda: self.getDirectory())

        self.preProcessTextures_btn.clicked.connect(
            lambda: self.preProcessTextures(str(self.dir_lbl.text())))

        self.processTextures_btn.clicked.connect(
            lambda: self.parseTextureToResize(str(self.dir_lbl.text()), str(
                self.textureSize_comboBox.currentText())))

        self.packTextures_btn.clicked.connect(
            lambda: self.parseTextureDirsToPack(str(self.dir_lbl.text())))

        self.aChannel_checkBox.toggled.connect(
            lambda: self.toggleAlphaInput(self.aChannel_checkBox))

    def getDirectory(self):
        '''creates QFileDialog to find designated folder'''
        dlg = QtGui.QFileDialog.getExistingDirectory(
            None, 'Select a folder:', 'C:\\Users\\desktop', QtGui.QFileDialog.ShowDirsOnly)

        self.dir_lbl.setText(dlg)
        if dlg:
            self.textureTools_tab_widget.setEnabled(True)
            print self.dir_lbl.text()
        else:
            self.textureTools_tab_widget.setEnabled(False)
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
        '''
        # Author: A.Polino
        # https://code.activestate.com/recipes/577514-chek-if-a-number-is-a-power-of-two/
        '''

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
                        base_file, self.extensions = os.path.splitext(filename)
                        print base_file
                        print type(base_file)
                        OriginalFileNamePath = os.path.join(dirpath, filename)
                        print OriginalFileNamePath
                        NewFileNamePath = os.path.join(
                            dirpath, base_file + '.tga')
                        # in order to successfully rename a file 
                        # the file needs to be joined with the path
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

    def parseTextureToResize(self, psPath, targetSize):

        # counterer for number of files
        countTileable = 0

        # list used to collect images larger than targetSize
        largerThanTargetSize = []

        alreadySizedTextures = []

        notPowerOf2Texture = []

        notSquareTexture = []

        # walk through directory, sub directories, and files
        for (dirname, dirs, files) in os.walk(psPath):

            #print dirname
            #print dirs
            #print files
            if dirs:
                # iterate over list of subdirectories
                for d in dirs:
                    path = os.path.join(dirname, d)
                    
                    # precautionary check to ensure is valid path
                    if os.path.isdir(path):

                        largerThanTargetSize, alreadySizedTextures, notPowerOf2Texture, notSquareTexture \
                        = self.analyzeTexturesToResize( \
                        path, largerThanTargetSize, alreadySizedTextures, notPowerOf2Texture, notSquareTexture, targetSize)

            else:
                largerThanTargetSize, alreadySizedTextures, notPowerOf2Texture, notSquareTexture \
                = self.analyzeTexturesToResize( \
                dirname, largerThanTargetSize, alreadySizedTextures, notPowerOf2Texture, notSquareTexture, targetSize)

        if largerThanTargetSize:

            self.textureResize(largerThanTargetSize, targetSize)

            combinedLargerTexturesHeader = 'Resized Textures:' + '\n'
            combinedLargerTextures = '\n'.join(str(x) for x in largerThanTargetSize)
        else:
            combinedLargerTexturesHeader = 'Resized Textures' + '\n'
            combinedLargerTextures = 'No Textures found larger than ' + self.textureSize_comboBox.currentText() + ' were found'

        combinedAlreadyResizedTexturesHeader = '\n\n' + \
            'Already ReSized and Skipped:' + '\n'
        alreadyResizedTextures = '\n'.join(
            str(x) for x in alreadySizedTextures)

        combinedNotPowerOf2TexturesHeader = '\n\n' + 'Not Power of 2 Textures' + '\n'
        notPowerOf2Textures = '\n'.join(
            str(x) for x in notPowerOf2Texture)

        combinedNotSquareTextureHeader = '\n\n' + 'Not Square Textures' + '\n'
        notSquareTextures = '\n'.join(
            str(x) for x in notSquareTexture)

        outputString = combinedLargerTexturesHeader + \
            combinedLargerTextures + combinedAlreadyResizedTexturesHeader + \
            alreadyResizedTextures + combinedNotPowerOf2TexturesHeader + \
            notPowerOf2Textures + combinedNotSquareTextureHeader + \
            notSquareTextures

        self.popupOkWindow(outputString)

        largerThanTargetSize = []

        alreadySizedTextures = []

        notPowerOf2Texture = []

        notSquareTexture = []

    def analyzeTexturesToResize(self, path, largerThanTargetSize, alreadySizedTextures, notPowerOf2Texture, notSquareTexture, targetSize):

        # iterate over entries in found subdirectory
        for x in os.listdir(path):

            file = os.path.join(path, x)

            # check if x during iteration is a file
            if os.path.isfile(file):

                # check if file extension exists in extension list
                if x.lower().endswith(self.extensions):

                    # s variable used to iterate over self.textureSizes tuple
                    # x current element being processed in directory
                    if any(str(s) in x for s in self.textureSizes):
                        alreadySizedTextures.append(x)
                    else:

                        # use PIL package to create Image object
                        with Image.open(file) as im:
                            sizeOfImage = im.size

                        # sizeOfImage returns tuple (width, height)
                        # check that image is square by comparing width and height
                        if sizeOfImage[0] == sizeOfImage[1]:

                            # if width/height are equal, use either value to check if power of 2
                            if self.is_power2(sizeOfImage[0]):
                                print x + ' - ' + '{0}'.format(sizeOfImage)
                                # countTileable = countTileable + 1

                                if sizeOfImage[0] < targetSize:
                                    # testPrint = testPrint + imagePath + '\n'
                                    largerThanTargetSize.append(file)
                            else:
                                notPowerOf2Texture.append(x)
                        else:
                            notSquareTexture.append(x)
        
        return largerThanTargetSize, alreadySizedTextures, notPowerOf2Texture, notSquareTexture

    def textureResize(self, listToResize, targetSize):
        '''
        Logic to control Photoshop, resize textures, and output a new texture
        and save the new size in the file name
        '''

        osVersion = self.checkWindowsVersion()
        print self.checkPhotoshopVersion()
        psApp = self.launchPhotoshop(osVersion)

        version = psApp.Version
        print version

        print version
        print psApp.path

        for x in listToResize:
            test = psApp.Open(x)

            psApp.Application.ActiveDocument

            test.resizeImage(int(targetSize), int(targetSize))

            filename, file_extension = os.path.splitext(x)

            newFileName = filename + '_' + \
                str(targetSize) + file_extension
            self.saveTGA(osVersion, psApp, newFileName)

            # close original version without saving
            test.Close(2)

        self.closePS('Close Photoshop?', psApp)

    def parseTextureDirsToPack(self, psPath):
        '''
        Based on user input, code will determine whether to iterate
        through folders or across files to find textures to pack
        '''

        # counterer for number of files
        countTileable = 0

        # check to prevent execution with entred output suffix
        if self.packedTexture_le.text():

            # precautionary check to make sure that there is an entry in the RGB channels
            if self.rChannel_le.text() and self.gChannel_le.text() and self.bChannel_le.text():

                # walk through directory, sub directories, and files
                for (dirname, dirs, files) in os.walk(psPath):

                    # print dirname

                    # true if function input directory has subfolders
                    # else iterate over files
                    if dirs:
                        # print 'dirs'

                        # iterate over list of directories found by os.walk
                        for d in dirs:
                            # print d

                            # join iterated directories with parent directory
                            path = os.path.join(dirname, d)

                            self.packTextures(path)

                    else:
                        self.packTextures(dirname)
            else:
                self.popupOkWindow('No Suffix Entred for all RGB Channels')
        else:
            self.popupOkWindow('No Suffix for Packed Texture')

    def packTextures(self, path):
        '''
        Logic used to control Photoshop and copy flattened textures
        into RGBA channels of a new texture
        '''

        # variables used to store paths of textures that match desired suffixes
        # variables are initilized to empty strings
        rTexture = ''
        gTexture = ''
        bTexture = ''
        aTexture = ''

        # iterate over entries in a directory
        for x in os.listdir(path):

            # store current path
            currentPath = os.path.join(os.path.join(path, x))

            # check if entry in directory is a file
            if os.path.isfile(currentPath):

                # check if file extension exists in extension list
                if x.lower().endswith(self.extensions):

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
                    # should run only once within each directory if a texture is found for RGB at least
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

        # after using photoshop, prompt and ask user if they are done using photoshop
        self.closePS('Close Photoshop?', psApp)

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
        popupWindow.setText(str(message))
        popupWindow.setStandardButtons(QtGui.QMessageBox.Ok)

        popupWindow.exec_()

    def closePS(self, message, psApp):
        popupWindow = QtGui.QMessageBox()

        popupWindow.setText(str(message))
        popupWindow.setStandardButtons(QtGui.QMessageBox.Yes | QtGui.QMessageBox.No)
        popupWindow.setWindowFlags(QtCore.Qt.Popup)

        result = popupWindow.exec_()

        if result == QtGui.QMessageBox.Yes:
            psApp.Quit()

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