#!/usr/bin/env python3
# -*- Mode: python; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4 -*- */
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

import logging
import argparse
import os
import tempfile
import subprocess
import re

from shutil import copyfile

def convert_to_absolute_win_path(path):
    win_path = subprocess.check_output(["cygpath", "-w", path]).strip().decode("utf-8")
    logging.info("Converted %s to Windows path %s" % (path, win_path))
    return win_path

def check_needed_files_in_path():
    logging.info("Checking that all needed executables are in the PATH")
    necessary_executables = ["msidb.exe", "msimsp.exe"]
    for executable in necessary_executables:
        try:
            logging.info("Checking %s" % executable)
            subprocess.call(executable)
        except FileNotFoundError:
            logging.error("The mandatory executable %s could not be found in PATH. Please ensure that the MSI tools are found in the path" % executable)
            raise

def install_msi_file(msi_file):
    logging.info("Installing MSI file: %s" % msi_file)
    return "some_\\tpath"

def extract_all_tables_from_pcpfile(fullpcpfilename, localmspdir):
    command = ["msidb.exe", "-d", convert_to_absolute_win_path(fullpcpfilename), "-f", convert_to_absolute_win_path(localmspdir), "-e", "*"]
    try:
        subprocess.check_call(command)
    except:
        logging.error("Could not extract tables from %s with command %s" % (fullpcpfilename, command))
        raise

def check_and_save_tables(tablelist, localmspdir):
    for table in tablelist:
        filename = os.path.join(localmspdir, table + ".idt")
        filename_save = filename + ".sav"
        if not os.path.exists(filename):
            logging.error("Could not find mandatory IDT file: %s" % filename)
            raise FileNotFoundError("could not find %s" % filename)
        copyfile(filename, filename_save)

def generate_msp_file_name(localmspdir):
    return os.path.join(localmspdir, "LibreOffice.msp")

def get_guid():
    output = subprocess.check_output("uuidgen.exe").strip().decode("utf-8").upper()
    return output

def change_properties_table(localmspdir, mspfilename):
    logging.info("Changing content of table \"Properties.idt\"")
    filename = os.path.join(localmspdir, "Properties.idt")
    if not os.path.exists(filename):
        raise FileNotFoundError("Could not find %s" % filename)

    guid_string = "{" + get_guid() + "}"

    with open(filename, "r") as f:
        file_content = f.read()

    file_content = re.sub("PatchGUID\t.*", "PatchGUID\t%s" % guid_string, file_content)
    file_content = re.sub("PatchOutputPath\t.*", "PatchOutputPath\t%s" % convert_to_absolute_win_path(mspfilename).replace("\\t", "\\\\t"), file_content)

    with open(filename, "w") as f:
        f.write(file_content)

    logging.info("Successfully adapted the table \"Properties\"")

def change_target_images_table(localmspdir, olddatabase):
    logging.info("Changing content of table \"TargetImages\"")
    filename = os.path.join(localmspdir, "TargetImages.idt")
    if not os.path.exists(filename):
        raise FileNotFoundError("Could not find %s" % filename)

    with open(filename, "r") as f:
        file_content = f.read()

    file_content = "\n".join(file_content.split('\n', 3)[0:3])
    file_content = file_content + "\nT1\t%s\t\tU1\t1\t0x00000922\t1\n" % convert_to_absolute_win_path(olddatabase)

    with open(filename, "w") as f:
        f.write(file_content)

    logging.info("Successfully adapted the table \"TargetImages\"")

def change_upgraded_images_table(localmspdir, newdatabase):
    pass

def change_image_families_table(localmspdir):
    pass

def change_patch_metadata_table(localmspdir):
    pass

def change_patch_sequence_table(localmspdir):
    pass

def edit_tables(localmspdir, olddatabase, newdatabase, mspfilename):
    change_properties_table(localmspdir, mspfilename)
    change_target_images_table(localmspdir, olddatabase)
    change_upgraded_images_table(localmspdir, newdatabase)
    change_image_families_table(localmspdir)
    change_patch_metadata_table(localmspdir)
    change_patch_sequence_table(localmspdir)

def include_tables_into_pcpfile(fullpcpfilename, localmspdir, tablelist):
    pass

def execute_msimsp(fullpcpfilename, mspfilename, localmspdir):
    pass

def get_current_dir():
    dir_path = os.path.dirname(os.path.realpath(__file__))
    return dir_path

def get_local_pcp_file_path(pcp_file_name):
    return os.path.join(get_current_dir(), "pcp", pcp_file_name)

def create_msp_dir():
    temp_dir = tempfile.mkdtemp(prefix="msp_libreoffice")
    logging.info("Created MSP directory at %s" % temp_dir)
    return temp_dir

def create_msp_patch(old_msi_file, new_msi_file, sign = False):
    logging.info("*************************")
    logging.info("... creating msp file ...")
    logging.info("*************************")

    check_needed_files_in_path()

    olddatabase = install_msi_file(old_msi_file)
    newdatabase = install_msi_file(new_msi_file)

    pcpfilename = "libreoffice.pcp"

    localmspdir = create_msp_dir()
    fullpcpfilename = os.path.join(localmspdir, pcpfilename)

    # Copy pcp file into the msp directory
    copyfile(get_local_pcp_file_path(pcpfilename), fullpcpfilename)

    # Unpacking tables from pcp file
    extract_all_tables_from_pcpfile(fullpcpfilename, localmspdir)

    tablelist = ["Properties", "TargetImages", "UpgradedImages", "ImageFamilies", "PatchMetadata", "PatchSequence"]
    # Saving all tables
    check_and_save_tables(tablelist, localmspdir);

    # Setting the name of the new msp file
    msp_file_name = generate_msp_file_name(localmspdir)

    # Editing tables
    edit_tables(localmspdir, olddatabase, newdatabase, msp_file_name);

    # Adding edited tables into pcp file
    include_tables_into_pcpfile(fullpcpfilename, localmspdir, tablelist);

    execute_msimsp(fullpcpfilename, msp_file_name, localmspdir);

    if sign:
        # Handle signing the msp file
        logging.info("Signing the msp file")
        pass

    # some cleanup
    logging.info("Successfully created the msp file")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Generate MSP files from 2 MSI files")
    parser.add_argument('-l', '--log', default=False, help="Enable printing out the log output", action='store_const', const=True)
    parser.add_argument('-o', '--old', help="The path to the old MSI file", nargs=1, required=True)
    parser.add_argument('-n', '--new', help="The path to the new MSI file", nargs=1, required=True)
    parser.add_argument('-s', '--sign', default=False, help="Whether the generated msp file should be signed", action='store_const', const=True)

    args = parser.parse_args()
    if args.log:
        logging.getLogger().setLevel(logging.DEBUG)

    create_msp_patch(args.old, args.new, args.sign)

# vim:set shiftwidth=4 softtabstop=4 expandtab: */
