#!/usr/bin/python3

import os
import json
import jsons

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
        nail_stream = catalog_file.openstream('Nail')
        thumb_stream = catalog_file.openstream('Thumb')

        # READ THE CATEGORIES.
        self.unk1 = struct.unpack.uint32_le(category_stream)
        self.unk2 = struct.unpack.uint32_le(category_stream)
        category_count = struct.unpack.uint32_le(category_stream)
        self.unk3 = struct.unpack.uint32_le(category_stream)
        self.categories = []
        # The master category has the label "(All Categories)"
        # and includes all clips in this catalog.
        self.master_category = Category(category_stream)
        for index in range(category_count):
            category = Category(category_stream)
            self.categories.append(category)

        # READ THE FILE DEFINITION ENTRIES.
        thumb_stream.seek(0x190)
        self.entries = []
        entry = ClipEntry(thumb_stream)
        while entry._is_valid:
            self.entries.append(entry)
            entry = ClipEntry(thumb_stream)
        # After all the file definition entries have been read,
        # the rest of the thumb stream is junk. This contains
        # some interesting strings but isn't relevant to defining
        # cliparts, so we'll just throw it away.
        thumb_junk = thumb_stream.read()

    def export(self, export_directory_path):
        export_filename = os.path.basename(self.filepath) + '.json'
        export_filepath = os.path.join(export_directory_path, export_filename)
        with open(export_filepath, 'w') as json_file:
            self_as_dictionary = jsons.dump(self, strip_privates = True)
            self_as_json_string = json.dumps(self_as_dictionary, indent = 2)
            json_file.write(self_as_json_string)

## A clip art category in the Category stream.
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

## An entry in the thumb stream.
class ClipEntry:
    def __init__(self, stream):
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
            total_length = 640
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
