"""
# ==============================================================================
# !/usr/bin/env python
# title           :Pyotoshop.py
# description     :Script used to automate Photoshop operations across several files
# author          :Doug Halley
# date            :2017-12-16
# version         :11.0
# usage           :Standalone Python Application Executed by Pyotoshop.exe
# notes           :
# python_version  :2.7.14
# pyqt_version    :4.11.4
# ==============================================================================
"""

import os
import sys
import platform
import scandir
import comtypes.client

# from _ctypes import COMError
from PIL import Image

from PyQt4 import QtGui
from PyQt4 import QtCore

# Global Variables -------------------------------------------------------------
EXTENSIONS = ('.tga', '.png', '.jpg')
TEXTURE_SIZES = (4096, 2048, 1024, 512, 256, 128, 64)

class Main(QtGui.QMainWindow):  # pylint: disable = E1101
    '''
    The class that contains, defines, and creates the UI.
    '''

    def __init__(self, parent=None):
        """Initilizes the PyQt Interface.
        Keyword Arguments:
            parent {None} -- By having no parent, ui can be standalone
                                (default: {None})
        """

        super(Main, self).__init__(parent)
        self.create_ui()

    def create_ui(self):
        """Creates the PyQt Interface.
        All the PyQt logic needed is contained in this function,
        including the PyQt elements and connect functions.
        """

        # ==============================================================================
        # PYQT Widget Defintions
        # ==============================================================================

        # window title
        self.setWindowTitle('Pyotoshop')

        # sets ToolTip font for the UI
        QtGui.QToolTip.setFont(QtGui.QFont('SansSerif', 10)) # pylint: disable = E1101

        # main widget ------------------------------------------------------------------
        central_widget = QtGui.QWidget()  # pylint: disable = E1101
        central_widget.setLayout(QtGui.QVBoxLayout())  # pylint: disable = E1101

        # Search Directory widget ------------------------------------------------------
        input_dir_widget = QtGui.QVBoxLayout()  # pylint: disable = E1101

        self.add_directory = QtGui.QPushButton(  # pylint: disable = E1101
            'Choose Directory')
        self.add_directory.setToolTip(
            'Select Directory that contains textures to be processed')

        dir_name_lbl = QtGui.QLabel(  # pylint: disable = E1101
            'Selected Directory:')
        dir_name_lbl.setAlignment(QtCore.Qt.AlignCenter)  # pylint: disable = E1101

        self.dir_lbl = QtGui.QLabel('')  # pylint: disable = E1101
        self.dir_lbl.setAlignment(QtCore.Qt.AlignCenter)  # pylint: disable = E1101

        self.dir_lbl.setWordWrap(True)

        # parent widget for texture resizer elements -----------------------------------
        texture_resizer_widget = QtGui.QWidget()  # pylint: disable = E1101
        texture_resizer_widget.setLayout(
            QtGui.QVBoxLayout())  # pylint: disable = E1101

        # layout for combobox that contains varying texture sizes ----------------------
        texture_size_layout = QtGui.QHBoxLayout()  # pylint: disable = E1101

        texture_size_lbl = QtGui.QLabel(  # pylint: disable = E1101
            'Target Texture Size: ')

        # creates combobox for texture sizes
        self.texture_size_combobox = QtGui.QComboBox()  # pylint: disable = E1101

        for sizes in TEXTURE_SIZES:
            self.texture_size_combobox.addItem(
                QtCore.QString(str(sizes)))  # pylint: disable = E1101

        self.texture_size_combobox.setToolTip(
            'Available sizes to resize textures to a smaller resolution.')

        self.resize_textures_btn = QtGui.QPushButton(  # pylint: disable = E1101
            'Resize Textures')

        self.resize_textures_btn.setToolTip(
            'Resize available textures.')

        # Texture Packer widget to contain all elements need to run this tool ----------
        texture_packer_widget = QtGui.QWidget()  # pylint: disable = E1101
        texture_packer_widget.setLayout(QtGui.QVBoxLayout())  # pylint: disable = E1101

        # Input Channel form layout, child of texture_packer_widget --------------------
        input_channel_layout = QtGui.QFormLayout()  # pylint: disable = E1101

        channel_inputs_lbl = QtGui.QLabel(  # pylint: disable = E1101
            'Input Textures to Pack')
        channel_inputs_lbl.setAlignment(
            QtCore.Qt.AlignCenter)  # pylint: disable = E1101

        self.r_channel_le = QtGui.QLineEdit('')  # pylint: disable = E1101
        self.g_channel_le = QtGui.QLineEdit('')  # pylint: disable = E1101
        self.b_channel_le = QtGui.QLineEdit('')  # pylint: disable = E1101

        a_channel_layout = QtGui.QHBoxLayout()  # pylint: disable = E1101

        # QCheckbox is initialized at false
        self.a_channel_checkbox = QtGui.QCheckBox(  # pylint: disable = E1101
            'A Channel')
        self.a_channel_checkbox.setChecked(False)

        self.a_channel_checkbox.setToolTip(
            'Toggle whether or not an alpha channel is included in texture packing')

        # QLineEdit is initialized at false
        self.a_channel_le = QtGui.QLineEdit('')  # pylint: disable = E1101
        self.a_channel_le.setEnabled(False)

        channel_placeholder = 'Enter Suffix Here...'
        self.r_channel_le.setPlaceholderText(channel_placeholder)
        self.g_channel_le.setPlaceholderText(channel_placeholder)
        self.b_channel_le.setPlaceholderText(channel_placeholder)
        self.a_channel_le.setPlaceholderText(channel_placeholder)

        channel_tool_tip = 'Enter suffix for texture to search for packing'
        self.r_channel_le.setToolTip(channel_tool_tip)
        self.g_channel_le.setToolTip(channel_tool_tip)
        self.b_channel_le.setToolTip(channel_tool_tip)
        self.a_channel_le.setToolTip(channel_tool_tip)

        # output format layout declaration to contain string stating the output format -
        # child widgets of input_channel_layout ----------------------------------------
        output_texture_layout = QtGui.QHBoxLayout()  # pylint: disable = E1101

        self.output_format_lbl = QtGui.QLabel(  # pylint: disable = E1101
            '24bit .tga')

        output_texture_lbl = QtGui.QLabel(  # pylint: disable = E1101
            'Output Packed Texture - ' + self.output_format_lbl.text())

        output_texture_lbl.setAlignment(
            QtCore.Qt.AlignCenter)  # pylint: disable = E1101

        self.packed_texture_le = QtGui.QLineEdit('')  # pylint: disable = E1101

        self.pack_textures_btn = QtGui.QPushButton(  # pylint: disable = E1101
            'Pack Textures')

        self.packed_texture_le.setToolTip(
            'Enter suffix to add to the created packed texture')
        self.pack_textures_btn.setToolTip(
            'Create packed texture from textures with input suffixes')

        # ================================================================================
        # PYQT Widget Assignments
        # ================================================================================

        # Assigments for texture resizer widget ------------------------------------------
        texture_size_layout.addWidget(texture_size_lbl)
        texture_size_layout.addWidget(self.texture_size_combobox)

        # Assigments for texture packer parent widget ------------------------------------
        texture_resizer_widget.layout().addLayout(texture_size_layout)
        texture_resizer_widget.layout().addWidget(self.resize_textures_btn)

        # Assigments for texture packer childeren widgets --------------------------------
        input_dir_widget.addWidget(self.add_directory)
        input_dir_widget.addWidget(dir_name_lbl)
        input_dir_widget.addWidget(self.dir_lbl)

        a_channel_layout.addWidget(self.a_channel_checkbox)
        a_channel_layout.addWidget(self.a_channel_le)

        output_texture_layout.addWidget(output_texture_lbl)

        input_channel_layout.addRow(channel_inputs_lbl)

        input_channel_layout.addRow(
            QtCore.QString('R Channel'), self.r_channel_le)  # pylint: disable = E1101
        input_channel_layout.addRow(
            QtCore.QString('G Channel'), self.g_channel_le)  # pylint: disable = E1101
        input_channel_layout.addRow(
            QtCore.QString('B Channel'), self.b_channel_le)  # pylint: disable = E1101
        input_channel_layout.addRow(a_channel_layout)

        input_channel_layout.addRow(output_texture_layout)

        input_channel_layout.addRow(
            QtCore.QString('Packed Texture'), self.packed_texture_le)  # pylint: disable = E1101

        texture_packer_widget.layout().addLayout(input_channel_layout)
        texture_packer_widget.layout().addWidget(self.pack_textures_btn)

        # tab widget to contain all the subset tools -----------------------------------
        # child widgets of central_widget ----------------------------------------------
        self.texture_tools_tab_widget = QtGui.QTabWidget()  # pylint: disable = E1101
        self.texture_tools_tab_widget.setLayout(
            QtGui.QHBoxLayout())  # pylint: disable = E1101

        self.texture_tools_tab_widget.addTab(
            texture_packer_widget, 'Pack Textures')
        self.texture_tools_tab_widget.addTab(
            texture_resizer_widget, 'Resize Textures')

        self.texture_tools_tab_widget.setEnabled(False)

        # adds directiory widget and tool widgets to central widget ---------------------
        central_widget.layout().addLayout(input_dir_widget)
        central_widget.layout().addWidget(self.texture_tools_tab_widget)

        # sets central widget for PyQt window
        self.setCentralWidget(central_widget)

        self.setFixedSize(self.sizeHint())

        # ================================================================================
        # PYQT Execution Connections
        # ================================================================================

        # triggers for buttons
        self.add_directory.clicked.connect(
            lambda: self.get_directory())  # pylint: disable = W0108

        self.resize_textures_btn.clicked.connect(lambda: self.parse_texture_to_resize(
            str(self.dir_lbl.text()), str(self.texture_size_combobox.currentText())))

        self.pack_textures_btn.clicked.connect(
            lambda: self.parse_texture_dirs_to_pack(str(self.dir_lbl.text())))

        self.a_channel_checkbox.toggled.connect(
            lambda: self.toggle_alpha_input(self.a_channel_checkbox))

    def get_directory(self):
        """Create popup file browser and stores path.

        Creates QFileDialog to find and store designated folder
        """

        dlg = QtGui.QFileDialog.getExistingDirectory(  # pylint: disable = E1101
            None, 'Select a folder:', 'C:\\Users\\desktop', \
            QtGui.QFileDialog.ShowDirsOnly)  # pylint: disable = E1101

        self.dir_lbl.setText(dlg)

        if dlg:
            self.texture_tools_tab_widget.setEnabled(True)
            print self.dir_lbl.text()
        else:
            self.texture_tools_tab_widget.setEnabled(False)
            print self.dir_lbl.text()

    def toggle_alpha_input(self, checkbox):
        """Toggles alpha channel QLineEdit usablity.

        Checks if input checkbox is checked and determines if
        a_channel_le is enabled or disabled

        Arguments:
            checkbox {QCheckBox} -- If checkbox is checked
            set parameters of other Q-Objects
        """

        if checkbox.isChecked():
            self.a_channel_le.setEnabled(True)
            self.output_format_lbl.setText('32bit .tga')
        else:
            self.a_channel_le.setEnabled(False)
            self.output_format_lbl.setText('24bit .tga')

    @classmethod
    def scandir_to_dict(cls, path):
        """Use scandir walk function to parse directories.

        The scandir package was developed by Ben Hoyt and incorporated into Python 3.5.
        He wrote a blogpost about it here:
        http://benhoyt.com/writings/scandir/

        When developing with os.walk I noticed that there are extra iterations happening.
        They are not apparent because the conditional statements I set up only show what
        I want to see. Inherently, to my understanding, os.walk will iterate across each
        directory and each file in each directory.

        By contrast scandir makes a tuple for each directory.
        Tuple Indexes:
            [0] {string}- directory path
            [1] {list} - list of sub directories
            [2] {list} - list of files in the directory path

        I referred to these links to understand how to use scandir:
        https://www.blog.pythonlibrary.org/2016/01/26/python-101-how-to-traverse-a-directory/
        https://www.python.org/dev/peps/pep-0471/#os-scandir

        This format makes the results of scandir.walk easy to work with.
        I planned to send the results of scandir.walk into other functions
        and rather than remembering which index to work with, I decided to
        make a dictionary with easily indentifiable keys to make reading
        back my code easier.

        Arguments:
            path {string} -- root directory folder

        Returns:
            dictionary -- dicionary containing the results of scandir.walk
        """

        # list of dictionaries
        scandir_entries = []

        # walk through directory, sub directories, and files
        for entry in scandir.walk(path):
            temp = {"Directory": '', 'Sub_Dirs': [], 'Files': []}
            temp['Directory'] = entry[0]
            temp['Sub_Dirs'] = entry[1]
            temp['Files'] = entry[2]

            scandir_entries.append(temp)

        return scandir_entries

    def parse_texture_to_resize(self, path, target_size):
        """Parse through root directory and determine which actions to take.
        Parses input root and uses self.scandir_to_dict to get the directory,
        all subfolders and files into a dictionary format.
        Arguments:
            path {string} -- path to analyze
            target_size {int} -- designated size to resize textures to
        """

        # dictionary used to collect images larger than target_size
        texture_analysis_dict = {"Larger Textures": [], "Already Sized Textures": [], \
                                    "Not Power of 2": [], "Not Square": []}

        # use self.scandir_to_dict to walk across the root directory
        # and output the values to a dictionary
        scandir_list = self.scandir_to_dict(path)

        # iterate across list of dictionary that contains the
        # root directory, subfolders, and files
        for entry in scandir_list:

            # print str(entry)
            texture_analysis_dict = self.analyze_textures_to_resize(
                str(entry['Directory']), entry['Files'], texture_analysis_dict, target_size)

        # if this key's list has values, run function to resize these textures
        if texture_analysis_dict["Larger Textures"]:

            self.texture_resize(
                texture_analysis_dict["Larger Textures"], target_size)

            # larger_textures = '\n'.join(
            #     str(x) for x in texture_analysis_dict["Larger Textures"])
        # else:
        #     larger_textures = 'No Textures found larger than ' + \
        #         self.texture_size_combobox.currentText() + ' were found'

        # self.popup_ok_window(larger_textures)

        # reset dictionary values after above code is complete
        texture_analysis_dict["Larger Textures"] = []
        texture_analysis_dict["Already Sized Textures"] = []
        texture_analysis_dict["Not Power of 2"] = []
        texture_analysis_dict["Not Square"] = []

    def analyze_textures_to_resize(self, directory_path, dir_files, texture_dict, target_size):
        """Parse directories to find textures and resize them.
        This function does a preliminary scan of the directory given by the user.
        If texture files are found they are analyzed and their resolution is compared
        to the input resolution to determine the appopriate action to take.
        If textures are found to be larger than the input, photoshop is opened and the
        texture is resized and a duplicate is saved. The file names of the duplicate
        files are modified to reflect the texture size they have been scaled to.
        Arguments:
            directory_path {string} -- Input directory to analyze and parse through
            dir_files {list} -- Input list of files to iterate through
            texture_dict {dictionary} -- Dictionary used to store different scenarios
                                            and return the results of the analysis
            target_size {int} -- Input texture sized that files will be sized to
        Returns:
            dictionary -- After found textures are analyzed, they are stored in the
                            dictionary variable to be used later
        """

        # iterate over scandir_entries in found subdirectory
        for current_file in dir_files:

            current_file_path = os.path.join(directory_path, current_file)

            # check if file extension exists in extension list
            if current_file.lower().endswith(EXTENSIONS):
                # s variable used to iterate over TEXTURE_SIZES tuple
                # x current element being processed in directory
                if any(str(s) in current_file for s in TEXTURE_SIZES):
                    texture_dict["Already Sized Textures"].append(
                        current_file)
                else:
                    # use PIL package to create Image object
                    with Image.open(current_file_path) as image:
                        size_of_image = image.size

                    # size_of_image returns tuple (width, height)
                    # check that image is square by comparing width and height
                    if size_of_image[0] == size_of_image[1]:

                        # if width/height are equal,
                        # use either value to check if power of 2
                        if self.is_power2(size_of_image[0]):
                            # print current_file_path + ' - ' + '{0}'.format(size_of_image)
                            # count_tileable = count_tileable + 1

                            if size_of_image[0] < target_size:
                                # testPrint = testPrint + imagePath + '\n'
                                texture_dict["Larger Textures"].append(
                                    current_file_path)
                        else:
                            texture_dict["Not Power of 2"].append(current_file)
                    else:
                        texture_dict["Not Square"].append(current_file)

        return texture_dict

    def texture_resize(self, list_to_resize, target_size):
        """Resize and export process textures.
        Logic to control Photoshop, resize textures, and save as a new texture
        and include the new size in the file name
        Arguments:
            list_to_resize {list} -- List of textures designated to be resized
            target_size {int} -- Input texture sized that files will be sized to
        """

        os_version = self.check_windows_version()
        # print self.check_photoshop_version()
        ps_app = self.launch_photoshop(os_version)

        progdialog = self.popup_progress_window(
            'Resizing Textures', len(list_to_resize))

        # Logic used to determine current version of
        # Photoshop in use, currently unnecessary
        # version = ps_app.Version
        # print version

        # print version
        # print ps_app.path

        for texture in list_to_resize:
            # open texture file in Photoshop
            current_ps_doc = ps_app.Open(texture)

            current_index = list_to_resize.index(texture)

            progdialog.setValue(current_index)

            progdialog.setLabelText("Resizing %s..." \
                % list_to_resize[current_index])

            # incase Photoshop was already open, make current
            # document the active document
            ps_app.Application.ActiveDocument  # pylint: disable = W0104

            # call the Photoshop resize operation
            current_ps_doc.resizeImage(int(target_size), int(target_size))

            # split the extension from the texture path
            file_name, file_extension = os.path.splitext(texture)

            # by splitting the extension, the new image size can be
            # appended to a new string and that is combined with the
            # extension
            new_file_name = file_name + '_' + \
                str(target_size) + file_extension

            # call function to Save As the resized texture
            self.save_as(os_version, ps_app, current_ps_doc, new_file_name)

            # close original version without saving
            current_ps_doc.Close(2)

        progdialog.setValue(len(list_to_resize))

        progdialog.close()

        # launch popup to ask user if they are done with photoshop
        self.close_photoshop(ps_app)

    def parse_texture_dirs_to_pack(self, path):
        """Parse through root directory and determine which actions to take.
        Parses input root and uses self.scandir_to_dict to get the directory,
        all subfolders and files into a dictionary format.
        Arguments:
            path {string} -- path to analyze
            target_size {int} -- designated size to resize textures to
        """

        texture_analysis_dict = {"Red": [], "Green": [], "Blue": [], "Alpha": []}

        # check to prevent execution without entered output suffix
        if self.packed_texture_le.text():

            # precautionary check to make sure that there is an entry in the RGB channels
            if self.r_channel_le.text() and self.g_channel_le.text() and \
                self.b_channel_le.text():

                # use self.scandir_to_dict to walk across the root directory
                # and output the values to a dictionary
                scandir_list = self.scandir_to_dict(path)

                # iterate across list of dictionary that contains the
                # root directory, subfolders, and files
                for entry in scandir_list:

                    texture_analysis_dict = self.analyze_textures_to_pack(
                        str(entry['Directory']), entry['Files'], texture_analysis_dict)

                # if this key's list has values, run function to pack these textures
                if texture_analysis_dict["Red"] and \
                        texture_analysis_dict["Blue"] and texture_analysis_dict["Blue"]:

                    # print str(texture_analysis_dict)
                    self.pack_textures(texture_analysis_dict)

            else:
                self.popup_ok_window('No Suffix Entred for all RGB Channels')
        else:
            self.popup_ok_window('No Suffix for Packed Texture')

    def analyze_textures_to_pack(self, directory_path, dir_files, texture_dict):
        """Analyze files within a directory and determine if the directory contains
        the designate textures to pack.
        Textures within the directory are iterated over and compared to the inputs
        of the RGB QLineEdits and if there are matches, store those paths in another
        dictionary to be used to pack the textures.
        Arguments:
            directory_path {string} -- Input directory to analyze and parse through
            dir_files {list} -- Input list of files to iterate through
            texture_dict {dictionary} -- Dictionary used to store different scenarios
                                            and return the results of the analysis
        Returns:
            dictionary -- After found textures are analyzed, they are stored in the
                            dictionary variable to be used later for texture packing
        """

        # variables used to store paths of textures that match desired suffixes
        # variables are initilized to empty strings
        r_texture = ''
        g_texture = ''
        b_texture = ''
        a_texture = ''

        # iterate over scandir_entries in a directory
        for current_file in dir_files:

            # store current path by joining directory and the current file name
            current_file_path = os.path.join(directory_path, current_file)

            # check if file extension exists in extension list
            if current_file.lower().endswith(EXTENSIONS):

                # series of conditions to check if the current file matches any
                # of the QLineEdit channel inputs.
                # secondary condition to check if the temporary string variable
                # is empty.
                if str(self.r_channel_le.text()) in current_file:
                    if not r_texture:
                        print current_file
                        r_texture = current_file_path

                if str(self.g_channel_le.text()) in current_file:
                    if not g_texture:
                        print current_file
                        g_texture = current_file_path

                if str(self.b_channel_le.text()) in current_file:
                    if not b_texture:
                        print current_file
                        b_texture = current_file_path

                # similarly to RGB checks, will only check if A QLineEdit is enabled
                if self.a_channel_le.isEnabled() and self.a_channel_le.text():
                    if str(self.a_channel_le.text()) in current_file:
                        if not a_texture:
                            print current_file
                            a_texture = current_file_path

            # print 'current file ' + current_file
            # print 'current index ' + str(dir_files.index(current_file))
            # print 'length ' + str(len(dir_files) - 1)

            # if last iteration/element of directory files to search through
            if dir_files.index(current_file) == len(dir_files) - 1:

                # check if a value for R, G, and B is found to continue
                # should run only once within each directory
                # if a texture is found for RGB at least
                if r_texture and g_texture and b_texture:

                    texture_dict['Red'].append(r_texture)
                    texture_dict['Green'].append(g_texture)
                    texture_dict['Blue'].append(b_texture)
                    texture_dict['Alpha'].append(a_texture)

                # clear variables on last iteration
                r_texture = ''
                g_texture = ''
                b_texture = ''
                a_texture = ''

        # print str(texture_dict)
        return texture_dict

    def pack_textures(self, scandir_entry):
        """Logic used to control Photoshop and copy flattened textures
            into RGBA channels of a new texture.
        Arguments:
            scandir_entry {dictionary} -- input scandir generated dictionary to
                                        iterate over.
        """

        # get version of Windows
        os_version = self.check_windows_version()

        # open Photoshop
        ps_app = self.launch_photoshop(os_version)

        progdialog = self.popup_progress_window(
            'Packing Textures', len(scandir_entry['Red']))

        # iterate over scandir_entries in a directory
        # since the length of scandir_entry['Red'], ['Blue'], ['Green'],
        # and ['Alpha'] are all the same, one of them can be iterated over
        # and the index number just needs to be matched for each other channel
        for current_file in scandir_entry['Red']:

            current_index = scandir_entry['Red'].index(current_file)

            progdialog.setValue(current_index)

            progdialog.setLabelText("Packing %s..." \
                % scandir_entry['Red'][current_index])

            # open texture matching designated suffix to be used for R Channel
            r_doc = ps_app.Open(scandir_entry['Red'][current_index])

            # get width and height of texture
            doc_width = r_doc.width
            doc_height = r_doc.height

            # selec and  copy contents of the layer in focus
            r_doc.selection.selectAll()
            r_doc.activeLayer.Copy()

            # use height and width variables to create new texture with same dimentions
            blank_doc = ps_app.Documents.Add(
                doc_width, doc_height, 72, 'new_document', 2, 1, 1)

            # blank_doc.channels['Red'] - equivalent to calling channel by name
            # activeChannels must receive an array
            blank_doc.activeChannels = [blank_doc.channels['Red']]
            blank_doc.Paste()

            progdialog.setLabelText("Packing %s..." \
                % scandir_entry['Green'][current_index])

            # follows same flow as what was done for R Channel
            g_doc = ps_app.Open(scandir_entry['Green'][current_index])
            g_doc.selection.selectAll()
            g_doc.activeLayer.Copy()

            ps_app.activeDocument = blank_doc
            blank_doc.activeChannels = [blank_doc.channels['Green']]
            blank_doc.Paste()

            progdialog.setLabelText("Packing %s..." \
                % scandir_entry['Blue'][current_index])

            # follows same flow as what was done for R and G Channels
            b_doc = ps_app.Open(scandir_entry['Blue'][current_index])
            b_doc.selection.selectAll()
            b_doc.activeLayer.Copy()

            ps_app.activeDocument = blank_doc
            blank_doc.activeChannels = [blank_doc.channels['Blue']]
            blank_doc.Paste()

            # close original textures without saving
            r_doc.Close(2)
            g_doc.Close(2)
            b_doc.Close(2)

            # based on earlier A Channel checks
            # should only proceed if A Channel was desired
            if scandir_entry['Alpha'][current_index]:

                progdialog.setLabelText("Packing %s..." \
                    % scandir_entry['Alpha'][current_index])

                # follows same flow as what was done for R, G and B Channels
                a_doc = ps_app.Open(scandir_entry['Alpha'][current_index])
                a_doc.selection.selectAll()
                a_doc.activeLayer.Copy()

                ps_app.activeDocument = blank_doc
                blank_doc.channels.add()
                # blank_doc.Name = 'Alpha 1'
                # blank_doc.Kind = 2  # = PsChannelType.psMaskedAreaAlphaChannel
                blank_doc.Paste()

                a_doc.Close(2)

            split_path, split_path_file_name = os.path.split(
                current_file)

            file_name, file_ext = os.path.splitext(split_path_file_name)

            # gets first element of split from '_' based on naming convention at BBA
            split_file_name = file_name.split('_')

            if len(split_file_name) > 1:
                split_file_name.pop()
                split_file_name = '_'.join(split_file_name)
            else:
                split_file_name = split_file_name[0]

            new_file_name = str(split_file_name) + \
                str(self.packed_texture_le.text()) + file_ext

            new_file_name_path = os.path.join(split_path, new_file_name)

            # if there is an alpha input be sure to export TGA with alpha option on
            if scandir_entry['Alpha'][current_index]:
                self.save_tga(os_version, ps_app, new_file_name_path, True)
            else:
                self.save_tga(os_version, ps_app, new_file_name_path)

            blank_doc.Close(2)

        progdialog.setValue(len(scandir_entry['Red']))

        progdialog.close()

        # after using photoshop, prompt and ask user if they are done using photoshop
        self.close_photoshop(ps_app)

    def check_windows_version(self):
        """Uses the platform package to determine the version of Windows.
        Returns:
            string -- Returns Windows OS version
        """

        current_platform = platform.platform()

        if 'Windows' in current_platform:
            #splits windows version based on dashes
            split_platform_name = current_platform.split('-')
            #returns windows version number
            return split_platform_name[1]
        else:
            self.popup_ok_window('Untested OS. Tool only works on Windows')

    def check_photoshop_version(self):
        """Determine version of Photoshop installed.
        Returns:
            string -- Returns Photoshop version
        """

        #default Photoshop install path
        if os.path.isdir('C:\\Program Files\\Adobe\\'):

            # get list of folders in default Photoshop install path
            listdir = os.listdir('C:\\Program Files\\Adobe\\')

            # Searches Adobe directory to determine to see
            # if a version of Photoshop is installed
            found_items = [x for x in listdir if 'Photoshop' in x]

            if found_items:

                # check how many versions of Photoshop are installed
                if len(found_items) == 1:
                    found_photoshop = str(found_items[0])

                    # check if a CC version of Photoshop is installed
                    if 'CC' in found_photoshop:

                        # remove spaces from version of Photoshop
                        split_photoshop = found_photoshop.split(' ')

                        # get last element from split list
                        ver_number = split_photoshop[-1]

                        # create empty string to store value and testing
                        version = ''

                        # version 14
                        if ver_number == 'CC':
                            version = '70'
                        elif int(ver_number):
                            # version 15
                            if ver_number == '2014':
                                version = '80'
                            # version 16
                            elif ver_number == '2015':
                                version = '90'
                            # version 17
                            elif ver_number == '2016':
                                version = '100'
                            # version 18
                            elif ver_number == '2017':
                                version = '110'
                            # version 19
                            elif ver_number == '2018':
                                version = '120'

                    else:
                        if version.startswith('12'):
                            version = '12'
                        elif version.startswith('13'):
                            version = '13'
                        else:
                            self.popup_ok_window('Non CC Version of PS')

                    if version:
                        return version
                    else:
                        self.popup_ok_window(
                            'Error getting installed Photoshop Version')
                else:
                    self.popup_ok_window(
                        'Multiple Versions of Photoshop installed')
            else:
                self.popup_ok_window('Photoshop not installed')
        else:
            self.popup_ok_window(
                'Adobe Software not installed in default directory')

    def launch_photoshop(self, os_ver):
        """Function used to launch Photoshop.
        Creates a com object that is used to launch and control Photoshop.
        Currently the com object is set to dynamic, unsure what this entails
        but it has allowed the script to work with various Photoshop versions.
        Photoshop versions tested with 2018, 2017, CS6
        Arguments:
            os_ver {string} -- Windows OS version
        Returns:
            com_object -- Returns instance of Photoshop object.
        """

        if os_ver == '10':
            ps_app = comtypes.client.CreateObject(
                'Photoshop.Application', dynamic=True)
            ps_app.Visible = True

            # Set the default unit to pixels!
            ps_app.Preferences.RulerUnits = 1

            return ps_app
        else:
            self.popup_ok_window(
                'Error with determining OS Version to launch Photoshop')

    def resize_results_popup(self, texture_dict, target_size):
        """Generates popup to show results of resize texture function.
        Collects contents of resize texture dictionary, combines into a string, and
        generates a popup witht his string.
        Arguments:
            texture_dict {dictionary} -- Dictionary used to store different scenarios
                                            and return the results of the analysis
            target_size {int} -- Input texture sized that files will be sized to
        """

        if texture_dict["Larger Textures"]:

            self.texture_resize(texture_dict["Larger Textures"], target_size)

            larger_textures_header = 'Resized Textures:' + '\n'
            larger_textures = '\n'.join(
                str(x) for x in texture_dict["Larger Textures"])
        else:
            larger_textures_header = 'Resized Textures' + '\n'
            larger_textures = 'No Textures found larger than ' + \
                self.texture_size_combobox.currentText() + ' were found'

        already_resized_textures_header = '\n\n' + \
            'Already ReSized and Skipped:' + '\n'
        already_resized_textures = '\n'.join(
            str(x) for x in texture_dict["Already Sized Textures"])

        not_power_of_2_textures_header = '\n\n' + 'Not Power of 2 Textures' + '\n'
        not_power_of_2_textures = '\n'.join(
            str(x) for x in texture_dict["Not Power of 2"])

        not_square_texture_header = '\n\n' + 'Not Square Textures' + '\n'
        not_square_textures = '\n'.join(
            str(x) for x in texture_dict["Not Square"])

        output_string = larger_textures_header + \
            larger_textures + already_resized_textures_header + \
            already_resized_textures + not_power_of_2_textures_header + \
            not_power_of_2_textures + not_square_texture_header + \
            not_square_textures

        self.popup_ok_window(output_string)

    def save_as(self, os_ver, ps_app, ps_doc, file_name):
        """Runs Save As Photoshop operation to save resized texture as a
            duplicate file.
        Similarly to com_object needed to launch and control Photoshop, a
        com_object is needed. So a com_object is created to contain the appopriate
        save options to save the file.
        Arguments:
            os_ver {string} -- Current version of Windows being used.
            ps_app {com_object} -- Gets current Photoshop instance
            ps_doc {com_object} -- The active photoshop document.
            file_name {string} -- File name for the tga file to be generated.
        """

        # check if currently on Win10
        if os_ver == '10':

            # split the extension from the texture path
            file_name_split, file_extension = os.path.splitext(file_name)

            if file_extension == '.tga':

                # get channel count of document
                # if there are less than 3 channels then save without alpha
                if ps_doc.channels.count > 3:
                    self.save_tga(os_ver, ps_app, file_name, True)
                else:
                    self.save_tga(os_ver, ps_app, file_name)
            else:
                if file_extension == '.jpg':
                    # creates com object for tga save operation
                    save_options = comtypes.client.CreateObject(
                        'Photoshop.JPEGSaveOptions', dynamic=True)
                elif file_extension == '.png':
                    # creates com object for tga save operation
                    save_options = comtypes.client.CreateObject(
                        'Photoshop.PNGSaveOptions', dynamic=True)

                # Photoshop saves texture as tga
                ps_app.ActiveDocument.SaveAs(file_name, save_options, True)

    @classmethod
    def save_tga(cls, os_ver, ps_app, tga_file, alpha_channel=False):
        """Runs Save As Photoshop operation to save resized texture as a
            duplicate file.
        Similarly to com_object needed to launch and control Photoshop, a
        com_object is needed. So a com_object is created to contain the
        attributes to create a tga file.
        Arguments:
            os_ver {string} -- Current version of Windows being used.
            ps_doc {com_object} -- The active photoshop document.
            tga_file {string} -- File name for the tga file to be generated.
        Keyword Arguments:
            alpha_channel {bool} -- Determines if alpha channel is included in
                                    save (default: {False})
        """

        # check if currently on Win10
        if os_ver == '10':

            # creates com object for tga save operation
            tga_save_options = comtypes.client.CreateObject(
                'Photoshop.TargaSaveOptions', dynamic=True)
            tga_save_options.Resolution = 24
            tga_save_options.AlphaChannels = False
            tga_save_options.RLECompression = False

            # If designated to include alpha, set parameters to do so
            if alpha_channel:
                tga_save_options.Resolution = 32
                tga_save_options.AlphaChannels = True

            # Photoshop saves texture as tga
            ps_app.ActiveDocument.SaveAs(tga_file, tga_save_options, True)

    @classmethod
    def is_power2(cls, num):
        """Performs calculation to determine if input
        number is a power of 2.
        Author: A.Polino
        https://code.activestate.com/recipes/577514-chek-if-a-number-is-a-power-of-two/
        Arguments:
            num {int} -- Input value to check
        Returns:
            boolean -- One line calculation and returns True or False
        """

        # states if a number is a power of two
        return num != 0 and ((num & (num - 1)) == 0)

    @classmethod
    def popup_detailed_ok_window(cls, message):
        """Generic popup window with an OK button and displays message.
        Generates QMessageBox with OK button.
        Used for a detailed notification.
        Arguments:
            message {string} -- string to be generated in detailed popup
        """

        popup_window = QtGui.QMessageBox()  # pylint: disable = E1101

        popup_window.setText('Textures Found')
        popup_window.setDetailedText(str(message))
        popup_window.setStandardButtons(
            QtGui.QMessageBox.Ok)  # pylint: disable = E1101

        popup_window.exec_()

    @classmethod
    def popup_progress_window(cls, window_title, progress_length):
        """Popup QProgressDialog to display operation progress.

        The progress_length parameter is the length of the input list or dictionary and
        while looping through the list or dictionary, the index is used to show the current
        progression.

        Arguments:
            window_title {str} -- Input string name that will appear on the
                QProgressDialog as the title of the dialog window
            progress_length {int} -- Input integer that sets the size of the QProgressDialog

        Returns:
            QProgressDialog -- Returns the QProgressDialog so that it can be accessed elsewhere.
        """

        progdialog = QtGui.QProgressDialog(  # pylint: disable = E1101
            "", "Cancel", 0, progress_length)

        progdialog.setWindowTitle(window_title)
        # progdialog.setWindowModality(
        #     QtCore.Qt.WindowModal)  # pylint: disable = E1101
        progdialog.setWindowFlags(
            QtCore.Qt.WindowStaysOnTopHint)  # pylint: disable = E1101

        progdialog.show()

        return progdialog

    @classmethod
    def popup_ok_window(cls, message):
        """Generic popup window with an OK button and displays message
        Generates QMessageBox with OK button.
        Used as a simple notification.
        Arguments:
            message {string} -- string to be generated in popup
        """

        popup_window = QtGui.QMessageBox()  # pylint: disable = E1101
        popup_window.setText(str(message))
        popup_window.setStandardButtons(
            QtGui.QMessageBox.Ok)  # pylint: disable = E1101

        popup_window.exec_()

    @classmethod
    def close_photoshop(cls, ps_app):
        """Popup to ask user if they would want to close Photoshop
        Generates QMessageBox with yes and no buttons.
        If yes button is clicked close Photoshop
        Arguments:
            ps_app {com_object} -- Gets current Photoshop instance
        """

        popup_window = QtGui.QMessageBox()  # pylint: disable = E1101

        popup_window.setText('Close Photoshop?')
        popup_window.setStandardButtons(
            QtGui.QMessageBox.Yes | QtGui.QMessageBox.No)  # pylint: disable = E1101
        popup_window.setWindowFlags(QtCore.Qt.Popup)  # pylint: disable = E1101

        result = popup_window.exec_()

        if result == QtGui.QMessageBox.Yes:  # pylint: disable = E1101
            ps_app.Quit()

if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)  # pylint: disable = E1101
    my_widget = Main()
    my_widget.show()
sys.exit(app.exec_())
