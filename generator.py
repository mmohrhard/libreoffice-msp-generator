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
import datetime

from shutil import copyfile

def convert_to_absolute_win_path(path):
    win_path = subprocess.check_output(["cygpath", "-w", path]).strip().decode("utf-8")
    logging.debug("Converted %s to Windows path %s" % (path, win_path))
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
    temp_dir = tempfile.mkdtemp(prefix="msp_libreoffice_msi_dir")
    logging.info("Installing MSI file: %s into %s" % (msi_file, temp_dir))
    command = ["msiexec.exe", "/a", msi_file, "TARGET_DIR=%s" % convert_to_absolute_win_path(temp_dir), "/qn", "/l*v", "admin_install.log"]
    try:
        subprocess.check_call(command)
    except:
        logging.error("Could not successfully install %s into %s with command %s. Error log can be found in \"admin_install.log\"" %(msi_file, temp_dir, " ".join(command)))
        raise
    return temp_dir

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
    logging.info("Changing content of table \"UpgradedImages\"")
    filename = os.path.join(localmspdir, "UpgradedImages.idt")
    if not os.path.exists(filename):
        raise FileNotFoundError("Could not find %s" % filename)

    with open(filename, "r") as f:
        file_content = f.read()

    file_list = file_content.split('\n', 4)
    file_content = "\n".join(file_list[0:3])
    upgraded = "U1"
    patchmsipath = ""
    symbolpath = ""
    family = "22334455"

    # TODO: moggi: handle the case of an existing line 3
    logging.warn("Currently ignoring existing content for UpgradedImages: %s" % file_list[3])

    file_content = file_content + "\n%s\t%s\t%s\t%s\t%s\n" % (upgraded, convert_to_absolute_win_path(newdatabase), patchmsipath, symbolpath, family)

    with open(filename, "w") as f:
        f.write(file_content)

    logging.info("Successfully adapted the table \"UpgradedImages\"")

def change_image_families_table(localmspdir):
    logging.info("Changing content of table \"ImageFamilies\"")
    filename = os.path.join(localmspdir, "ImageFamilies.idt")
    if not os.path.exists(filename):
        raise FileNotFoundError("Could not find %s" % filename)

    with open(filename, "r") as f:
        file_content = f.read()

    file_list = file_content.split('\n', 4)
    file_content = "\n".join(file_list[0:3])
    family = "22334455"
    media_src_propname = "MediaSrcPropName"
    media_disk_id = "2"
    file_sequence_start = "0" # TODO: moggi: this needs to be calculated
    disk_prompt = ""
    volume_label = ""

    # TODO: moggi: handle the case of an existing line 3
    logging.warn("Currently ignoring existing content for ImageFamilies: %s" % file_list[3])

    file_content = file_content + "\n%s\t%s\t%s\t%s\t%s\t%s\n" % (family, media_src_propname, media_disk_id, file_sequence_start, disk_prompt, volume_label)

    with open(filename, "w") as f:
        f.write(file_content)

    logging.info("Successfully adapted the table \"ImageFamilies\"")

def get_patch_sequence():
    package_version = os.environ.get('LIBO_PACKAGEVERSION').strip()
    match = re.match('(\d+)\.(\d+)\.(\d+)\.(\d+)', package_version) # $major.$minor.$micro.$patch
    print(match.groups())
    if len(match.groups()) != 4:
        raise Exception("The 'LIBO_PACKAGEVERSION' environment variable needs to have the form '$major.$minor.$micro.$patch")
    return package_version

def change_patch_metadata_table(localmspdir):
    logging.info("Changing content of table \"PatchMetadata\"")
    filename = os.path.join(localmspdir, "PatchMetadata.idt")
    if not os.path.exists(filename):
        raise FileNotFoundError("Could not find %s" % filename)

    with open(filename, "r") as f:
        file_content = f.read()

    data = {}
    data['Classification'] = ('', os.environ.get('LIBO_SERVICEPACK', 'Hotfix'))
    data['AllowRemoval'] = ('', os.environ.get('LIBO_ALLOWREMOVAL', '1'))
    data['CreationTimeUTC'] = ('', datetime.datetime.utcnow().strftime("%m/%d/%Y %H:%M"))

    if data['Classification'][1] == 'Hotfix':
        service_pack = False
    elif data['Classification'][1] == 'ServicePack':
        service_pack = True
    else:
        logging.error("Unknown Classification type: \"%s\"; Only \"Hotfix\" and \"ServicePack\" allowed." % data['Classification'][1])
        raise Exception("Unknown Classification type: %s" % data['Classification'][1])

    product_name = os.environ.get('LIBO_PRODUCTNAME', 'LibreOffice')
    product_version = os.environ.get('LIBO_PRODUCTVERSION')

    base = product_name + " " + product_version

    data['TargetProductName'] = ('', product_name)
    data['ManufacturerName'] = ('', os.environ.get('LIBO_VENDOR', 'LibreOffice'))
    patch_sequence_value = get_patch_sequence()

    build_id = os.environ.get('LIBO_BUILDID', '123')

    if service_pack:
        windows_level_value = os.environ.get('LIBO_PATCHLEVEL', '0')
        name_and_descr = base + " ServicePack " + windows_patch_level + " " + patch_sequence_value + " Build: " + build_id
        data['DisplayName'] = ('', name_and_descr.strip())
        data['Description'] = ('', name_and_descr.strip())
    else:
        display_addon = os.environ.get('LIBO_PATCH_DISPLAY_ADDON', '')
        name_and_descr = base + " HotFix " + display_addon + " " + patch_sequence_value + " Build: " + build_id
        data['DisplayName'] = ('', name_and_descr.strip())
        data['Description'] = ('', name_and_descr.strip())

    split_file_content = file_content.split("\n", 3)

    for line in split_file_content[3].split("\n"):
        match = re.match("^\s*(.*?)\t(.*?)\t(.*?)\s*$", line)
        if match is None:
            continue
        if len(match.groups()) != 3:
            continue

        if match.groups()[1] in data:
            data[match.groups()[1]] = (match.groups()[0], data[match.groups()[1]][1]) # set the company from the original file
        else:
            data[match.groups()[1]] = (match.groups()[0], match.groups()[2])

    with open(filename, "w") as f:
        f.write(split_file_content[0])
        f.write("\n")
        f.write(split_file_content[1])
        f.write("\n")
        f.write(split_file_content[2])
        f.write("\n")
        for key, value in data.items():
            line = "%s\t%s\t%s\n" % (value[0], key, value[1])
            f.write(line)

    logging.info("Successfully changed content of table \"PatchMetadata\"")

def get_super_sede():
    service_pack = os.environ.get('LIBO_SERVICEPACK', 'Hotfix')
    data = {'Hotfix' : '0', 'ServicePack' : '1'}
    return data[service_pack]

def change_patch_sequence_table(localmspdir):
    logging.info("Changing content of table \"PatchSequence\"")
    filename = os.path.join(localmspdir, "PatchSequence.idt")
    if not os.path.exists(filename):
        raise FileNotFoundError("Could not find %s" % filename)

    with open(filename, "r") as f:
        file_content = f.read()

    split_file_content = file_content.split("\n", 4)
    file_content = "\n".join(split_file_content[0:3])

    patch_family = "SO"
    target = ""
    patch_sequence = get_patch_sequence()
    super_sede = get_super_sede()
    if len(split_file_content) > 3:
        line = split_file_content[3].split("\n")[0]
        match = re.match("^\s*(.*?)\t(.*?)\t(.*?)\t(.*?)\s*", line)
        if match:
            patch_family = match.groups()[0]
            target = match.groups()[1]

    file_content = file_content + "\n%s\t%s\t%s\t%s\n" % (patch_family, target, patch_sequence, super_sede)
    with open(filename, "w") as f:
        f.write(file_content)

def edit_tables(localmspdir, olddatabase, newdatabase, mspfilename):
    change_properties_table(localmspdir, mspfilename)
    change_target_images_table(localmspdir, olddatabase)
    change_upgraded_images_table(localmspdir, newdatabase)
    change_image_families_table(localmspdir)
    change_patch_metadata_table(localmspdir)
    change_patch_sequence_table(localmspdir)

def include_tables_into_pcpfile(fullpcpfilename, localmspdir, tablelist):
    for table in tablelist:
        if len(table) > 8:
            old_name = os.path.join(localmspdir, table + ".idt")
            new_name = os.path.join(localmspdir, table[0:8] + ".idt")
            logging.info("Copying table from old name \"%s\" to \"%s\" to comply with 8+3 naming rules" % (old_name, new_name))
            copyfile(old_name, new_name)

    for table in tablelist:
        command = ["msidb.exe", "-d", convert_to_absolute_win_path(fullpcpfilename), "-f", convert_to_absolute_win_path(localmspdir), "-i", table]
        try:
            subprocess.check_call(command)
        except:
            logging.error("Failed to execute: %s" % ("\n".join(command)))
            raise

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

    create_msp_patch(args.old[0], args.new[0], args.sign)

# vim:set shiftwidth=4 softtabstop=4 expandtab: */
