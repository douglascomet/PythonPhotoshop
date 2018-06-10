import os
import sys
import platform
import traceback
import scandir
import comtypes.client
from PIL import Image

EXTENSIONS = ('.tga', '.png', '.jpg')
TEXTURE_SIZES = (4096, 2048, 1024, 512, 256, 128, 64)


class Pyotoshop(object):
    def __init__(self):
        super(Pyotoshop, self).__init__()

    def scandir_to_dict(cls, path):
        """Converts the tuples created from scandir walk to dictionaries.

        Tuple Indexes:
            [0] {string}- directory path
            [1] {list} - list of sub directories
            [2] {list} - list of files in the directory path

        Arguments:
            path (string): root directory folder

        Returns:
            dictionary -- dicionary containing the results of scandir.walk
        """

        # list of dictionaries
        scandir_entries = []

        # walk through directory, sub directories, and files
        for entry in scandir.walk(path):
            temp = {}
            temp['Directory'] = entry[0]
            temp['Sub_Dirs'] = entry[1]
            temp['Files'] = entry[2]

            scandir_entries.append(temp)

        return scandir_entries

    def parse_texture_to_resize(self, path):
        """Parse through root directory and determine which actions to take.
        Parses input root and uses self.scandir_to_dict to get the directory,
        all subfolders and files into a dictionary format.

        Arguments:
            path (string): path to analyze
        """

        # dictionary used to collect images larger than target_size
        texture_analysis_dict = {
            'Larger Textures': [],
            'Already Sized Textures': [],
            'Not Power of 2': [],
            'Not Square': []}

        # use self.scandir_to_dict to walk across the root directory
        # and output the values to a dictionary
        scandir_list = self.scandir_to_dict(path)

        progress_dialog = self.popup_progress_window(
            'Finding Textures to Resize', len(scandir_list))

        # iterate across list of dictionary that contains the
        # root directory, subfolders, and files
        for index, entry in enumerate(scandir_list):

            if progress_dialog.wasCanceled():
                texture_analysis_dict.clear()
                break

            # print str(entry)
            texture_analysis_dict = self.analyze_textures_to_resize(
                str(entry['Directory']), entry['Files'], texture_analysis_dict)

            progress_dialog.setValue(index)

            progress_dialog.setLabelText(
                'Searching for Textures in {0}...'.format(str(entry['Directory'])))

        progress_dialog.setValue(len(scandir_list))

        progress_dialog.close()

        print texture_analysis_dict['Larger Textures']
        # if this key's list has values, run function to resize these textures
        if texture_analysis_dict['Larger Textures']:

            self.texture_resize(texture_analysis_dict['Larger Textures'])

            # larger_textures = '\n'.join(
            #     str(x) for x in texture_analysis_dict['Larger Textures'])
        else:
            # larger_textures = 'No Textures found larger than ' + \
            #     self.texture_size_combobox.currentText() + ' were found'

            self.popup_ok_window('No Textures found to resize')

        # reset dictionary values after above code is complete
        texture_analysis_dict['Larger Textures'] = []
        texture_analysis_dict['Already Sized Textures'] = []
        texture_analysis_dict['Not Power of 2'] = []
        texture_analysis_dict['Not Square'] = []

    def analyze_textures_to_resize(self, directory_path, dir_files, texture_dict):
        """Parse directories to find textures and resize them.
        This function does a preliminary scan of the directory
        given by the user.
        If texture files are found they are analyzed and their
        resolution is compared to the input resolution to determine
        the appopriate action to take.
        If textures are found to be larger than the input, photoshop is
        opened and the texture is resized and a duplicate is saved. The file
        names of the duplicate files are modified to reflect the texture size
        they have been scaled to.

        Arguments:
            directory_path (string): Input directory to analyze and
                parse through
            dir_files (list): Input list of files to iterate through
            texture_dict (dictionary): Dictionary used to store
                different scenarios and return the results of the analysis

        Returns:
            dictionary -- After found textures are analyzed,
                they are stored in the dictionary variable to be used later
        """

        # iterate over scandir_entries in found subdirectory
        for current_file in dir_files:

            current_file_path = os.path.join(directory_path, current_file)

            # check if file extension exists in extension list
            if current_file.lower().endswith(EXTENSIONS):

                # s variable used to iterate over TEXTURE_SIZES tuple
                # x current element being processed in directory
                if any(str(s) in current_file for s in TEXTURE_SIZES):
                    texture_dict['Already Sized Textures'].append(current_file)
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
                            # print current_file_path + ' - ' + \
                            #   '{0}'.format(size_of_image)
                            # count_tileable = count_tileable + 1

                            if int(size_of_image[0]) > int(
                                    self.target_texture_size_combobox.currentText()):
                                # testPrint = testPrint + imagePath + '\n'
                                texture_dict['Larger Textures'].append(current_file_path)
                            else:
                                print current_file_path + ' - ' + '{0}'.format(size_of_image)
                        else:
                            texture_dict['Not Power of 2'].append(current_file)
                    else:
                        texture_dict['Not Square'].append(current_file)

        return texture_dict

    def texture_resize(self, list_to_resize):
        """Resize and export process textures.
        Logic to control Photoshop, resize textures, and save as a new texture
        and include the new size in the file name

        Arguments:
            list_to_resize (list): List of textures designated to be resized
        """

        ps_app = self.launch_photoshop()

        progress_dialog = self.popup_progress_window('Resizing Textures', len(list_to_resize))

        for texture_path in list_to_resize:

            if progress_dialog.wasCanceled():
                self.popup_ok_window('Search Canceled')
                break

            path_name = os.path.dirname(os.path.abspath(texture_path))

            # open texture file in Photoshop
            current_ps_doc = ps_app.Open(texture_path)

            current_index = list_to_resize.index(texture_path)

            progress_dialog.setValue(current_index)

            progress_dialog.setLabelText('Resizing Textures in {0}...'.format(path_name))

            # incase Photoshop was already open, make current
            # document the active document
            ps_app.Application.ActiveDocument  # pylint: disable = W0104

            target_resolution = int(self.target_texture_size_combobox.currentText())

            # call the Photoshop resize operation
            current_ps_doc.resizeImage(target_resolution, target_resolution)

            new_file_name = self.new_file_name(texture_path, True)

            # call function to Save As the resized texture
            self.save_as(ps_app, current_ps_doc, new_file_name)

            # close original version without saving
            current_ps_doc.Close(2)

        progress_dialog.setValue(len(list_to_resize))

        progress_dialog.close()

        # launch popup to ask user if they are done with photoshop
        self.close_photoshop('Completed Texture Resizing!', ps_app)

    def parse_texture_dirs_to_pack(self, path):
        """Parse through root directory and determine which actions to take.
        Parses input root and uses self.scandir_to_dict to get the directory,
        all subfolders and files into a dictionary format.

        Arguments:
            path (string): path to analyze
        """

        texture_analysis_dict = {
            'Red': [],
            'Green': [],
            'Blue': [],
            'Alpha': []}

        # check to prevent execution without entered output suffix
        if self.packed_texture_le.text():

            # precautionary check to make sure that there is an
            # entry in the RGB channels
            if self.r_channel_le.text() and self.g_channel_le.text() and \
                    self.b_channel_le.text():

                # use self.scandir_to_dict to walk across the root directory
                # and output the values to a dictionary
                scandir_list = self.scandir_to_dict(path)

                progress_dialog = self.popup_progress_window(
                    'Finding Textures to Pack', len(scandir_list))

                # iterate across list of dictionary that contains the
                # root directory, subfolders, and files
                for index, entry in enumerate(scandir_list):

                    if progress_dialog.wasCanceled():
                        texture_analysis_dict.clear()
                        self.popup_ok_window('Search Canceled')
                        break

                    texture_analysis_dict = self.analyze_textures_to_pack(
                        str(entry['Directory']), entry['Files'],
                        texture_analysis_dict)

                    progress_dialog.setValue(index)

                    progress_dialog.setLabelText(
                        'Searching for Textures in {0}...'.format(
                            str(entry['Directory'])))

                progress_dialog.setValue(len(scandir_list))

                progress_dialog.close()

                if texture_analysis_dict:
                    # if this key's list has values, run function to
                    # pack these textures
                    if texture_analysis_dict['Red'] and \
                        texture_analysis_dict['Blue'] and \
                            texture_analysis_dict['Blue']:

                        # print str(texture_analysis_dict)
                        self.pack_textures(texture_analysis_dict)

                else:
                    self.popup_ok_window('No textures were found to pack')
            else:
                self.popup_ok_window('No Suffix Entred for all RGB Channels')
        else:
            self.popup_ok_window('No Suffix for Packed Texture')

    def analyze_textures_to_pack(self, directory_path, dir_files, texture_dict):
        """Analyze files within a directory and determine if the
            directory contains the designate textures to pack.

        Textures within the directory are iterated over and compared to the
        inputs of the RGB QLineEdits and if there are matches, store those
        paths in another dictionary to be used to pack the textures.

        Arguments:
            directory_path (string): Input directory to analyze and
                                        parse through
            dir_files (list): Input list of files to iterate through
            texture_dict (dictionary): Dictionary used to
                                        store different scenarios and
                                        return the results of the analysis
        Returns:
            dictionary -- After found textures are analyzed, they are stored
                            in the dictionary variable to be used
                                later for texture packing
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

                # similarly to RGB checks, will only check if a
                # QLineEdit is enabled
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
            scandir_entry (dictionary): input scandir generated dictionary to
                                        iterate over.
        """

        # open Photoshop
        ps_app = self.launch_photoshop()

        progress_dialog = self.popup_progress_window(
            'Packing Textures', len(scandir_entry['Red']))

        # iterate over scandir_entries in a directory
        # since the length of scandir_entry['Red'], ['Blue'], ['Green'],
        # and ['Alpha'] are all the same, one of them can be iterated over
        # and the index number just needs to be matched for each other channel
        for current_file in scandir_entry['Red']:

            if progress_dialog.wasCanceled():
                self.popup_ok_window('Search Canceled')
                break

            new_file_name_path = self.new_file_name(current_file)

            current_index = scandir_entry['Red'].index(current_file)

            progress_dialog.setValue(current_index)

            progress_dialog.setLabelText(
                'Packing {0}...'.format(new_file_name_path))

            # open texture matching designated suffix to be used
            # for R Channel
            r_doc = ps_app.Open(scandir_entry['Red'][current_index])

            # get width and height of texture
            doc_width = r_doc.width
            doc_height = r_doc.height

            # selec and  copy contents of the layer in focus
            r_doc.selection.selectAll()
            r_doc.activeLayer.Copy()

            # use height and width variables to create new texture
            # with same resolution
            blank_doc = ps_app.Documents.Add(
                doc_width, doc_height, 72, 'new_document', 2, 1, 1)

            # blank_doc.channels['Red'] - equivalent to calling channel
            # by name
            # activeChannels must receive an array
            blank_doc.activeChannels = [blank_doc.channels['Red']]
            blank_doc.Paste()

            # follows same flow as what was done for R Channel
            g_doc = ps_app.Open(scandir_entry['Green'][current_index])
            g_doc.selection.selectAll()
            g_doc.activeLayer.Copy()

            ps_app.activeDocument = blank_doc
            blank_doc.activeChannels = [blank_doc.channels['Green']]
            blank_doc.Paste()

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

                # follows same flow as what was done for R, G and B Channels
                a_doc = ps_app.Open(scandir_entry['Alpha'][current_index])
                a_doc.selection.selectAll()
                a_doc.activeLayer.Copy()

                ps_app.activeDocument = blank_doc
                blank_doc.channels.add()
                # blank_doc.Name = 'Alpha 1'
                # blank_doc.Kind = 2
                # = PsChannelType.psMaskedAreaAlphaChannel
                blank_doc.Paste()

                a_doc.Close(2)

            # if there is an alpha input be sure to export TGA with
            # alpha option on
            if scandir_entry['Alpha'][current_index]:
                self.save_tga(ps_app, new_file_name_path, True)
            else:
                self.save_tga(ps_app, new_file_name_path)

            blank_doc.Close(2)

        progress_dialog.setValue(len(scandir_entry['Red']))

        progress_dialog.close()

        # after using photoshop, prompt and ask user if they are done
        # using photoshop
        self.close_photoshop('Completed Texture Packing!', ps_app)

    def new_file_name(self, file_path, resize=False):
        """Since assigning a new file name for both texture packing and
        texture resizing follow similar operations, the functions were
        combined.

        Arguments:
            file_path (str): full path of texture

        Keyword Arguments:
            resize (bool): Toggle to be able switch between
            texture packing or texture resizing (default: {False})

        Returns:
            str -- Returns updated path name
        """

        split_path, split_path_file_name = os.path.split(file_path)

        file_name, file_ext = os.path.splitext(split_path_file_name)

        if resize:
            # split the extension from the texture path
            file_name, file_extension = os.path.splitext(file_path)

            # by splitting the extension, the new image size can be
            # appended to a new string and that is combined with the
            # extension
            new_file_name = file_name + '_' + str(
                self.target_texture_size_combobox.currentText()
                ) + file_extension

            return new_file_name
        else:
            # gets first element of split from '_' based on naming
            # convention at BBA
            split_file_name = file_name.split('_')

            if len(split_file_name) > 1:
                split_file_name.pop()
                split_file_name = '_'.join(split_file_name)
            else:
                split_file_name = split_file_name[0]

            new_file_name = str(split_file_name) + \
                str(self.packed_texture_le.text()) + file_ext

            new_file_name_path = os.path.join(split_path, new_file_name)

            return new_file_name_path

    def check_windows_version(self):
        """Uses the platform package to determine the version of Windows.

        Returns:
            string -- Returns Windows OS version
        """

        current_platform = platform.platform()

        if 'Windows' in current_platform:
            # splits windows version based on dashes
            split_platform_name = current_platform.split('-')
            # returns windows version number
            return split_platform_name[1]
        else:
            self.popup_ok_window('Untested OS. Tool only works on Windows')

    def check_photoshop_version(self):
        """Determine version of Photoshop installed.

        Returns:
            string -- Returns Photoshop version
        """

        # default Photoshop install path
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

    def launch_photoshop(self):
        """Function used to launch Photoshop.

        Creates a com object that is used to launch and control Photoshop.
        Currently the com object is set to dynamic, unsure what this entails
        but it has allowed the script to work with various Photoshop versions.
        Photoshop versions tested with 2018, 2017, CS6

        Returns:
            com_object -- Returns instance of Photoshop object.
        """

        os_ver = self.check_windows_version()

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

    def resize_results_popup(self, texture_dict):
        """Generates popup to show results of resize texture function.
        Collects contents of resize texture dictionary, combines into a
        string, and generates a popup witht his string.

        Arguments:
            texture_dict (dictionary): Dictionary used to store different scenarios and return the
                                       results of the analysis

            target_size (int): Input texture sized that files will be sized to
        """

        if texture_dict['Larger Textures']:

            self.texture_resize(texture_dict['Larger Textures'])

            larger_textures_header = 'Resized Textures:' + '\n'
            larger_textures = '\n'.join(str(x) for x in texture_dict['Larger Textures'])
        else:
            larger_textures_header = 'Resized Textures' + '\n'
            larger_textures = 'No Textures found larger than ' + \
                str(self.target_texture_size_combobox.currentText()) + \
                ' were found'

        already_resized_textures_header = '\n\n' + 'Already ReSized and Skipped:' + '\n'
        already_resized_textures = '\n'.join(str(x) for x in texture_dict['Already Sized Textures'])

        not_power_of_2_textures_header = '\n\n' + 'Not Power of 2 Textures' + '\n'
        not_power_of_2_textures = '\n'.join(str(x) for x in texture_dict['Not Power of 2'])

        not_square_texture_header = '\n\n' + 'Not Square Textures' + '\n'
        not_square_textures = '\n'.join(str(x) for x in texture_dict['Not Square'])

        output_string = larger_textures_header + \
            larger_textures + already_resized_textures_header + \
            already_resized_textures + not_power_of_2_textures_header + \
            not_power_of_2_textures + not_square_texture_header + \
            not_square_textures

        self.popup_ok_window(output_string)

    def save_as(self, ps_app, ps_doc, file_name):
        """Runs Save As Photoshop operation to save resized texture as a
            duplicate file.

        Similarly to com_object needed to launch and control Photoshop, a
        com_object is needed. So a com_object is created to contain the
        appopriate save options to save the file.

        Arguments:
            ps_app (com_object): Gets current Photoshop instance
            ps_doc (com_object): The active photoshop document.
            file_name (string): File name for the tga file to be generated.
        """

        os_ver = self.check_windows_version()

        # check if currently on Win10
        if os_ver == '10':

            # split the extension from the texture path
            file_name_split, file_extension = os.path.splitext(file_name)

            if file_extension == '.tga':

                # get channel count of document
                # if there are less than 3 channels then save without alpha
                if ps_doc.channels.count > 3:
                    self.save_tga(ps_app, file_name, True)
                else:
                    self.save_tga(ps_app, file_name)

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

    def save_tga(self, ps_app, tga_file, alpha_channel=False):
        """Runs Save As Photoshop operation to save resized texture as a duplicate file.

        Similarly to com_object needed to launch and control Photoshop, a com_object is needed.
        So a com_object is created to contain the attributes to create a tga file.

        Arguments:
            ps_doc (com_object): The active photoshop document.
            tga_file (string): File name for the tga file to be generated.

        Keyword Arguments:
            alpha_channel (bool): Determines if alpha channel is included in save (default: {False})
        """

        os_ver = self.check_windows_version()

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
        """Performs calculation to determine if input number is a power of 2.

        Author: A.Polino
        https://code.activestate.com/recipes/577514-chek-if-a-number-is-a-power-of-two/

        Arguments:
            num (int): Input value to check

        Returns:
            boolean -- One line calculation and returns True or False
        """

        # states if a number is a power of two
        return num != 0 and ((num & (num - 1)) == 0)
