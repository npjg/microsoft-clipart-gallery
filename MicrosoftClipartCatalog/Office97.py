#!/usr/bin/python3

import os
import json
import jsons

from PIL import Image
import olefile
import self_documenting_struct as struct

## This has been tested with the English (US) version of
## Office 97. Other languages might require a change of 
## text encoding.
TEXT_ENCODING = 'latin-1'

## A CAG file, which contains all the metadata for this
## catalog. An OLE file that contains three streams:
##  - Category: A list of all the categories in this catalog.
##  - Nail: 88x88-pixel bitmap previews of each clip,
##          presented in the same order as the clips
##          in the "(All Categories)" category.
##  - Thumb: A list of each clipart file and the keywords
##           attached to it. 
class MicrosoftClipArt30Catalog:
    def __init__(self, catalog_filepath):
        # READ THE CATALOG FILE.
        self.filepath = catalog_filepath
        catalog_file = olefile.OleFileIO(catalog_filepath)
        category_stream = catalog_file.openstream('Category')
        # The "Nail" stream is the one that actually holds the thumbnails.
        nail_stream = catalog_file.openstream('Nail')
        # The "Thumb" stream actually holds metadata about each clip,
        # not thumbnails.
        thumb_stream = catalog_file.openstream('Thumb')

        # READ THE CATEGORIES.
        self.unk1 = struct.unpack.uint32_le(category_stream)
        self.unk2 = struct.unpack.uint32_le(category_stream)
        category_count = struct.unpack.uint32_le(category_stream)
        self.unk3 = struct.unpack.uint32_le(category_stream)
        self._categories = []
        for index in range(category_count):
            category = Category(category_stream)
            self._categories.append(category)

        # READ THE CLIPART DECLARATIONS.
        thumb_stream.seek(0x190)
        nail_stream.seek(0x800)
        self.clipart_declarations = []
        clipart_declaration = ClipartDeclaration(thumb_stream, nail_stream)
        while clipart_declaration._is_valid:
            self.clipart_declarations.append(clipart_declaration)
            clipart_declaration = ClipartDeclaration(thumb_stream, nail_stream)
        # After all the file definition entries have been read,
        # the rest of the thumb stream is junk. This contains
        # some interesting strings but isn't relevant to defining
        # cliparts, so we'll just throw it away.
        self._thumb_junk = thumb_stream.read()

        # PUT THE CLIPART IN THE PROPER CATEGORIES.
        for clip_id, clipart_declaration in zip(self.master_category.clip_ids, self.clipart_declarations):
            clipart_declaration.id = clip_id
            for category in self._categories:
                if clip_id in category.clip_ids:
                    clipart_declaration.categories.append(category.title)

    ## The master category has the ASCII label "(All Categories)" and includes the IDs of all clips in this catalog. 
    @property
    def master_category(self) -> 'Category':
        # However, it's not always clear which one this is, as sometimes there are multiple categories with this name
        # in ASCII that are only differentiated by some special characters at the end of the name. For example:
        #  - "(All Categories)\u0001\u0001"
        #  - "(All Categories)\u0001\u0002"
        # So we will treat the category that contains the same number of clip IDs as the number of clipart declarations
        # as the master category.
        for category in self._categories:
            if len(category.clip_ids) - 1 == len(self.clipart_declarations):
                return category
        raise ValueError('Master category not found. This Clip Art Catalog 3.0 file is probably corrupt.')

    ## Exports the data in this clipart catalog file to the filesystem.
    def export(self, export_directory_path):
        # EXPORT THE JSON.
        json_filename = os.path.basename(self.filepath) + '.json'
        json_filepath = os.path.join(export_directory_path, json_filename)
        with open(json_filepath, 'w') as json_file:
            self_as_dictionary = jsons.dump(self, strip_privates = True)
            self_as_json_string = json.dumps(self_as_dictionary, indent = 2)
            json_file.write(self_as_json_string)

        # EXPORT THE JUNK DATA.
        thumb_junk_filename = os.path.basename(self.filepath) + '.thumb_junk.dat'
        thumb_junk_filepath = os.path.join(export_directory_path, thumb_junk_filename)
        with open(thumb_junk_filepath, 'wb') as thumb_junk_file:
            thumb_junk_file.write(self._thumb_junk)

        # EXPORT THE THUMBNAILS.
        for clipart_declaration in self.clipart_declarations:
            thumbnail_filename = clipart_declaration.filename + '.thumbnail.bmp'
            print(thumbnail_filename)
            thumbnail_filepath = os.path.join(export_directory_path, thumbnail_filename)
            clipart_declaration._thumbnail.save(thumbnail_filepath)
            input()

## A clip art category, which contains references to clip art images by ID.
## Stored in the "Category" stream.
class Category:
    def __init__(self, stream):
        # READ THE TITLE.
        # TODO: Separate out the name from any binary ID that might
        # be present.
        self.title = struct.unpack.pascal_string(stream).decode(TEXT_ENCODING)
        if self.title == '\x00':
            # THIS IS THE LAST CATEGORY.
            # Attempting to read any further would read past the end 
            # of the stream.
            return

        # READ THE IDS OF THE CLIPS IN THIS CATEGORY.
        self.clip_ids = []
        total_clips_in_category = struct.unpack.uint16_le(stream)
        self.unk1 = struct.unpack.uint32_le(stream)
        for index in range(total_clips_in_category):
            clip_id = struct.unpack.uint32_le(stream)
            self.clip_ids.append(clip_id)

## Contains metadata about a single clipart image.
## Stored in the "Thumb" stream.
class ClipartDeclaration:
    def __init__(self, stream, thumbnail):
        # SET THE CATEGORY.
        # The ID and category will be assigned later.
        self._thumbnail: Image = Image.frombytes('P', (44, 88), thumbnail.read(44 * 88)).transpose(1)
        thumbnail.read(0xe0)
        self.id = None
        self.categories = []

        # READ THE TYPE.
        # I don't know why there are three different "types" of
        # entries, but the type of the entry determines how
        # long it is in the stream.
        entry_start = stream.tell()
        self._type = struct.unpack.uint32_le(stream)
        if self._type == 0x10:
            total_length = 0x320
        elif self._type == 0x20:
            total_length = 0x640
        elif self._type == 0x28:
            total_length = 0x640
        elif self._type == 0x30:
            total_length = 0x190
        elif self._type == 0xa0:
            total_length = 0x640
        elif self._type == 0x90:
            total_length = 0x320
        else:
            # STOP READING THIS ENTRY.
            # Reading an invalid type probably means we have read
            # all the data in the stream and all remaining data
            # is junk.
            self._is_valid = False
            return

        # FIND THE TOTAL LENGTH OF THIS ENTRY.
        entry_end = entry_start + total_length
        self._is_valid = True

        # READ THE STRINGS.
        self.filename = struct.unpack.pascal_string(stream).decode(TEXT_ENCODING)
        # If there is no subdirectory, this will just have a drive letter path
        # (usually "C:\").
        self.subdirectory = struct.unpack.pascal_string(stream).decode(TEXT_ENCODING)
        k = struct.unpack.pascal_string(stream)
        keywords_string = k.decode(TEXT_ENCODING)
        KEYWORD_SEPARATOR = ','
        self.keywords = keywords_string.split(KEYWORD_SEPARATOR)

        # READ THE REMAINING DATA.
        # It looks like this is just junk data that can be 
        # overwritten later when users add custom clips
        # to the catalog.
        remaining_data =  entry_end - stream.tell()
        # There is interesting stuff here, but it isn't relevant
        # to the actual clipart data so it will be thrown away.
        junk = stream.read(remaining_data)

class ThumbnailImage:
    pass
