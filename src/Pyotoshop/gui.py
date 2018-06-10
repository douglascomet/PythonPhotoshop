# -*- coding: utf-8 -*-

# Copyright (c) 2018 Doug Halley
#
# -*- coding:utf-8 -*-
# @Script: Pyotoshop.py
# @Author: Doug Halley
# @Email: douglascomet@gmail.com
# @Create At: 2018-05-31 22:05:50
# @Last Modified By: Doug Halley
# @Last Modified At: 2018-02-16 22:10:10
# @Description: This is description.
"""
# =============================================================================
# title           :Pyotoshop.py
# description     :Automate Photoshop operations across several files
# author          :Doug Halley
# date            :2018-01-06
# version         :15.0
# usage           :Standalone Python Application Executed by Pyotoshop.exe
# python_version  :2.7.14
# pyqt_version    :4.11.4
# =============================================================================
"""

import os
import sys
import platform
import scandir
import comtypes.client

from PIL import Image

from Qt import QtGui, QtCore, QtWidgets

# Global Variables ------------------------------------------------------------
EXTENSIONS = ('.tga', '.png', '.jpg')
TEXTURE_SIZES = (4096, 2048, 1024, 512, 256, 128, 64)


class Main(QtGui.QMainWindow):
    """
    The class that contains, defines, and creates the UI.
    """

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

        # =====================================================================
        # PYQT Widget Defintions
        # =====================================================================

        # window title
        self.setWindowTitle('Pyotoshop')

        # sets ToolTip font for the UI
        QtWidgets.QToolTip.setFont(QtWidgets.QFont('SansSerif', 10))

        # main widget ---------------------------------------------------------
        central_widget = QtWidgets.QWidget()
        central_widget.setLayout(QtWidgets.QVBoxLayout())

        # Search Directory QFormLayout-----------------------------------------
        add_dir_widget = QtWidgets.QFormLayout()

        # Resize texture button layout, child of central_widget ---------------
        add_directory_btn_layout = QtWidgets.QHBoxLayout()

        # Spacer used to place add_directory_btn in the center of the ui
        add_directory_btn_spacer = QtWidgets.QSpacerItem(0, 0)

        add_directory_btn = QtWidgets.QPushButton('Choose Directory')
        add_directory_btn.setToolTip('Select Directory that contains textures to be processed')
        add_directory_btn.setFixedWidth(150)

        dir_name_lbl = QtWidgets.QLabel('Selected Directory:')
        dir_name_lbl.setAlignment(QtCore.Qt.AlignCenter)

        directory_lbl = QtWidgets.QLabel('')
        directory_lbl.setAlignment(QtCore.Qt.AlignCenter)
        directory_lbl.setWordWrap(True)

        # parent widget for texture resizer elements,
        # child of texture_tools_tab_widget -----------------------------------
        texture_resizer_widget = QtWidgets.QWidget()
        texture_resizer_widget.setLayout(QtWidgets.QVBoxLayout())

        texture_resize_description_lbl = QtWidgets.QLabel(
            'Finds textures that are larger than the target size and ' +
            'resizes textures accordingly into a duplicate file.')
        texture_resize_description_lbl.setAlignment(QtCore.Qt.AlignCenter)
        texture_resize_description_lbl.setWordWrap(True)

        # layout for combobox that contains varying texture sizes
        # child of target_size_combobox_layout ------------------------
        target_texture_size_formlayout = QtWidgets.QFormLayout()
        target_texture_size_formlayout.setRowWrapPolicy(
            QtWidgets.QFormLayout.DontWrapRows)
        target_texture_size_formlayout.setFieldGrowthPolicy(
            QtWidgets.QFormLayout.FieldsStayAtSizeHint)
        target_texture_size_formlayout.setLabelAlignment(QtCore.Qt.AlignRight)
        target_texture_size_formlayout.setFormAlignment(QtCore.Qt.AlignVCenter)

        # Resize texture button layout, child of texture_resizer_widget -------
        target_size_combobox_layout = QtWidgets.QHBoxLayout()

        # Spacer used to place add_directory_btn in the center of the ui
        target_size_combobox_l_spacer = QtWidgets.QSpacerItem(30, 50)
        target_size_combobox_r_spacer = QtWidgets.QSpacerItem(10, 50)

        # creates combobox for texture sizes
        self.target_texture_size_combobox = QtWidgets.QComboBox()
        self.target_texture_size_combobox.setFixedWidth(50)
        self.target_texture_size_combobox.setToolTip(
            'Available sizes to resize textures to a smaller resolution.')

        for sizes in TEXTURE_SIZES:
            self.target_texture_size_combobox.addItem(str(sizes))

        # texture_resize_btn_layout, child of texture_resizer_widget ----------
        texture_resize_btn_layout = QtWidgets.QHBoxLayout()

        # Spacer used to place pack_textures_btn in the center of the ui
        resize_pack_btn_spacer = QtWidgets.QSpacerItem(10, 5)

        resize_textures_btn = QtWidgets.QPushButton('Resize Textures')
        resize_textures_btn.setToolTip('Resize available textures.')
        resize_textures_btn.setFixedWidth(150)

        # Texture Packer widget to contain all elements need to run this tool -
        texture_packer_widget = QtWidgets.QWidget()
        texture_packer_widget.setLayout(QtWidgets.QVBoxLayout())
        texture_packer_widget.layout().setAlignment(QtCore.Qt.AlignRight)

        # input_channel_formlayout, child of texture_packer_widget ------------
        input_channel_formlayout = QtWidgets.QFormLayout()
        input_channel_formlayout.setRowWrapPolicy(
            QtWidgets.QFormLayout.DontWrapRows)
        input_channel_formlayout.setFieldGrowthPolicy(
            QtWidgets.QFormLayout.FieldsStayAtSizeHint)
        input_channel_formlayout.setLabelAlignment(QtCore.Qt.AlignRight)
        input_channel_formlayout.setFormAlignment(QtCore.Qt.AlignVCenter)

        texture_pack_description_lbl = QtWidgets.QLabel(
            'Finds textures that having matching suffixes and ' +
            'packs them into a new texture.')
        texture_pack_description_lbl.setAlignment(QtCore.Qt.AlignCenter)
        texture_pack_description_lbl.setWordWrap(True)

        channel_inputs_lbl = QtWidgets.QLabel('- Enter Suffixes of Texture Files to Search for -')
        channel_inputs_lbl.setAlignment(QtCore.Qt.AlignCenter)

        self.r_channel_le = QtWidgets.QLineEdit('')
        self.g_channel_le = QtWidgets.QLineEdit('')
        self.b_channel_le = QtWidgets.QLineEdit('')

        # QCheckbox is initialized at false
        self.a_channel_checkbox = QtWidgets.QCheckBox('A Channel')
        self.a_channel_checkbox.setChecked(False)

        self.a_channel_checkbox.setToolTip(
            'Toggle whether or not an alpha ' + 'channel is included in texture packing')

        # QLineEdit is initialized at false
        self.a_channel_le = QtWidgets.QLineEdit('')
        self.a_channel_le.setEnabled(False)

        channel_placeholder = 'Enter Suffix Here...'
        self.r_channel_le.setPlaceholderText(channel_placeholder)
        self.g_channel_le.setPlaceholderText(channel_placeholder)
        self.b_channel_le.setPlaceholderText(channel_placeholder)
        self.a_channel_le.setPlaceholderText(channel_placeholder)

        channel_tool_tip = 'Enter suffix for texture to search'
        self.r_channel_le.setToolTip(channel_tool_tip)
        self.g_channel_le.setToolTip(channel_tool_tip)
        self.b_channel_le.setToolTip(channel_tool_tip)
        self.a_channel_le.setToolTip(channel_tool_tip)

        self.r_channel_le.setFixedWidth(110)
        self.g_channel_le.setFixedWidth(110)
        self.b_channel_le.setFixedWidth(110)
        self.a_channel_le.setFixedWidth(110)

        output_format_lbl = QtWidgets.QLabel('- Packed Texture Format - 24 bit .tga -')
        output_format_lbl.setAlignment(QtCore.Qt.AlignCenter)

        self.packed_texture_le = QtWidgets.QLineEdit('')
        self.packed_texture_le.setPlaceholderText(channel_placeholder)
        self.packed_texture_le.setToolTip(
            'Enter suffix to add to the created packed texture')
        self.packed_texture_le.setFixedWidth(110)

        # Output texture_pack_btn_layout,
        # child of input_channel_formlayout -----------------------------------
        texture_pack_btn_layout = QtWidgets.QHBoxLayout()

        # Spacer used to place pack_textures_btn in the center of the ui
        texture_pack_btn_spacer = QtWidgets.QSpacerItem(10, 5)

        pack_textures_btn = QtWidgets.QPushButton('Create Packed Textures')
        pack_textures_btn.setToolTip(
            'Pack texture that contain input suffixes into a new texture')
        pack_textures_btn.setFixedWidth(150)

        # =====================================================================
        # PYQT Widget Assignments
        # =====================================================================

        # Assigments for add_directory_btn_layout -----------------------------

        add_directory_btn_layout.addItem(add_directory_btn_spacer)
        add_directory_btn_layout.addWidget(add_directory_btn)

        add_dir_widget.addRow(add_directory_btn_layout)
        add_dir_widget.addRow(dir_name_lbl)

        # Assigments for texture_pack_btn_layout ------------------------------

        texture_pack_btn_layout.addItem(texture_pack_btn_spacer)
        texture_pack_btn_layout.addWidget(pack_textures_btn)

        # Assigments for input_channel_formlayout -----------------------------

        input_channel_formlayout.addRow(channel_inputs_lbl)
        input_channel_formlayout.addRow(
            QtWidgets.QLabel('R Channel'), self.r_channel_le)
        input_channel_formlayout.addRow(
            QtWidgets.QLabel('G Channel'), self.g_channel_le)
        input_channel_formlayout.addRow(
            QtWidgets.QLabel('B Channel'), self.b_channel_le)
        input_channel_formlayout.addRow(self.a_channel_checkbox,
                                        self.a_channel_le)
        input_channel_formlayout.addRow(output_format_lbl)
        input_channel_formlayout.addRow(
            QtWidgets.QLabel('Packed Texture'), self.packed_texture_le)

        input_channel_formlayout.addRow(texture_pack_btn_layout)

        # Assigments for texture_packer_widget --------------------------------

        texture_packer_widget.layout().addLayout(input_channel_formlayout)

        # tab widget to contain all the subset tools --------------------------
        # child widgets of central_widget -------------------------------------
        texture_tools_tab_widget = QtWidgets.QTabWidget()
        texture_tools_tab_widget.setLayout(QtWidgets.QHBoxLayout())

        texture_tools_tab_widget.addTab(texture_packer_widget, 'Pack Textures')
        texture_tools_tab_widget.addTab(texture_resizer_widget,
                                        'Resize Textures')
        texture_tools_tab_widget.setEnabled(False)

        # Assigments for target_texture_size_formlayout -----------------------
        target_texture_size_formlayout.addRow(
            QtWidgets.QLabel('Target Texture Size:'),
            self.target_texture_size_combobox)

        # Assigments for texture resizer QFormLayout with spacers -------------
        target_size_combobox_layout.addItem(target_size_combobox_l_spacer)

        target_size_combobox_layout.addLayout(target_texture_size_formlayout)

        target_size_combobox_layout.addItem(target_size_combobox_r_spacer)

        # Assigments for texture resize button layout -------------------------
        texture_resize_btn_layout.addItem(resize_pack_btn_spacer)
        texture_resize_btn_layout.addWidget(resize_textures_btn)

        # Assigments for texture packer parent widget -------------------------
        texture_resizer_widget.layout().addWidget(
            texture_resize_description_lbl)
        texture_resizer_widget.layout().addLayout(target_size_combobox_layout)
        texture_resizer_widget.layout().addLayout(texture_resize_btn_layout)

        # adds directiory widget and tab widgets to central widget ------------
        central_widget.layout().addLayout(add_dir_widget)
        central_widget.layout().addWidget(directory_lbl)
        central_widget.layout().addWidget(texture_tools_tab_widget)

        # sets central widget for PyQt window
        self.setCentralWidget(central_widget)
        self.setFixedSize(self.sizeHint())

        # =====================================================================
        # PYQT Execution Connections
        # =====================================================================

        # triggers for buttons
        add_directory_btn.clicked.connect(
            lambda: self.get_directory(directory_lbl, texture_tools_tab_widget))

        resize_textures_btn.clicked.connect(
            lambda: self.parse_texture_to_resize(str(directory_lbl.text())))

        pack_textures_btn.clicked.connect(
            lambda: self.parse_texture_dirs_to_pack(str(directory_lbl.text())))

        self.a_channel_checkbox.toggled.connect(
            lambda: self.toggle_alpha_input(
                self.a_channel_checkbox, self.a_channel_le, output_format_lbl
            )
        )

    @classmethod
    def get_directory(cls, directory_lbl, tab_widget):
        """Create popup file browser and stores path.

        Creates QFileDialog to find and store designated folder
        """

        dlg = QtWidgets.QFileDialog.getExistingDirectory(
            None, 'Select a folder:', 'C:\\Users\\desktop',
            QtWidgets.QFileDialog.ShowDirsOnly)

        directory_lbl.setText(dlg)

        if dlg:
            tab_widget.setEnabled(True)
            # print directory_lbl.text()
        else:
            tab_widget.setEnabled(False)
            # print directory_lbl.text()

    @classmethod
    def toggle_alpha_input(cls, checkbox, line_edit, label_text):
        """Toggles alpha channel QLineEdit usablity.

        Checks if input checkbox is checked and determines if
        a_channel_le is enabled or disabled

        Arguments:
            checkbox {QCheckBox} -- If checkbox is checked set parameters of other Q-Objects
            line_edit {QLineEdit} -- enabled or disabled based on checkbox parameter
            label_text {QLabel} -- text displays differently based on checkbox parameter
        """

        if checkbox.isChecked():
            line_edit.setEnabled(True)
            label_text.setText('- Packed Texture Format - 32 bit .tga -')
        else:
            line_edit.setEnabled(False)
            label_text.setText('- Packed Texture Format - 24 bit .tga -')

    @classmethod
    def popup_detailed_ok_window(cls, message):
        """Generic popup window with an OK button and displays message.
        Generates QMessageBox with OK button. Used for a detailed notification.

        Arguments:
            message {string} -- string to be generated in detailed popup
        """

        popup_window = QtWidgets.QMessageBox()

        popup_window.setText('Textures Found')
        popup_window.setDetailedText(str(message))
        popup_window.setStandardButtons(
            QtWidgets.QMessageBox.Ok)

        popup_window.exec_()

    @classmethod
    def popup_progress_window(cls, window_title, progress_length):
        """Popup QProgressDialog to display operation progress.

        The progress_length parameter is the length of the input list or
        dictionary and while looping through the list or dictionary, the index
        is used to show the current progression.

        Arguments:
            window_title {str} -- Input string name that will appear on the QProgressDialog
                                  as the title of the dialog window
            progress_length {int} -- Input integer that sets the size of the QProgressDialog

        Returns:
            QProgressDialog -- Returns the QProgressDialog so that it can be
                                accessed elsewhere.
        """

        # Creates the QProgressDialog Window
        progress_dialog = QtWidgets.QProgressDialog('', 'Cancel', 0, progress_length)

        # Sets window title
        progress_dialog.setWindowTitle(window_title)

        # Keeps the window on top of all other windows
        progress_dialog.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)

        # Show window
        progress_dialog.show()

        return progress_dialog

    @classmethod
    def popup_ok_window(cls, message):
        """Generic popup window with an OK button and displays message
        Generates QMessageBox with OK button. Used as a simple notification.

        Arguments:
            message {string} -- string to be generated in popup
        """

        popup_window = QtWidgets.QMessageBox()
        popup_window.setText(str(message))
        popup_window.setStandardButtons(QtWidgets.QMessageBox.Ok)

        popup_window.exec_()

    @classmethod
    def close_photoshop(cls, message, ps_app):
        """Popup to ask user if they would want to close Photoshop
        Generates QMessageBox with yes and no buttons.
        If yes button is clicked close Photoshop

        Arguments:
            ps_app {com_object} -- Gets current Photoshop instance
        """

        popup_window = QtWidgets.QMessageBox()

        popup_message = message + '\n' + 'Close Photoshop?'

        popup_window.setText(popup_message)
        popup_window.setStandardButtons(
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)

        popup_window.setWindowFlags(QtCore.Qt.Popup)

        result = popup_window.exec_()

        if result == QtWidgets.QMessageBox.Yes:
            ps_app.Quit()

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv) # create an application object
    my_widget = Main()
    my_widget.show()

sys.exit(app.exec_()) # informs environment that widget was destroyed
