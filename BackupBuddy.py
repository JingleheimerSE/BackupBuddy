#!/usr/bin/env python
###############################################################################################################################################################
# Script to backup certain system files.
###############################################################################################################################################################
import sys
import logging
import argparse
import os
import subprocess
import json
import shutil
from winreg import *
from getpass import getpass


###############################################################################################################################################################
# Configuration
###############################################################################################################################################################
SEVEN_ZIP_PATH = "C:\Program Files\\7-Zip\\7z.exe"


# Backup config
OPTIONAL_KEY = "optional"
FUNCTIONS_KEY = "functions"
BACKUP_FUNCTION_KEY = "backup"
RESTORE_FUNCTION_KEY = "restore"
ARCHIVE_KEY = "archive"
PATHS_KEY = "paths"
FILTERS_KEY = "filters"
APP_KEY ="app"
SOURCE_KEY = "source"

BACKUP_ITEMS = {
    "environment": {
        FUNCTIONS_KEY: {
            BACKUP_FUNCTION_KEY: "BackupEnvironment",
            RESTORE_FUNCTION_KEY: "RestoreEnvironment",
        },
    },
    "ssh": {
        ARCHIVE_KEY: {
            PATHS_KEY: [os.path.join(os.getenv('USERPROFILE'), ".ssh")]
        }
    },
    "git": {
        ARCHIVE_KEY: {
            PATHS_KEY: [os.path.join(os.getenv('USERPROFILE'), ".gitconfig")]
        },
        # Update from https://tortoisegit.org/download/
        APP_KEY: "https://download.tortoisegit.org/tgit/2.15.0.0/TortoiseGit-2.15.0.0-64bit.msi"  
    },
    "syncthing": {
        ARCHIVE_KEY: {
            PATHS_KEY: [os.path.join(os.getenv('LOCALAPPDATA'), "Syncthing")],
            FILTERS_KEY: ['-xr!*.log']
        },
        APP_KEY: "https://github.com/canton7/SyncTrayzor/releases/latest/download/SyncTrayzorSetup-x64.exe"
    },
    "vscode": {
        FUNCTIONS_KEY: {
            BACKUP_FUNCTION_KEY: "BackupVsCode",
            RESTORE_FUNCTION_KEY: "RestoreVsCode",
        },
        APP_KEY: "https://code.visualstudio.com/sha/download?build=stable&os=win32-x64-user"
    },
    "corsair-icue": {
        # From https://www.corsair.com/us/en/explorer/icue/
        APP_KEY: "https://www3.corsair.com/software/CUE_V5/public/modules/windows/installer/Install%20iCUE.exe"
    },
    "logitech-g-hub": {
        # From https://www.logitechg.com/en-us/innovation/g-hub.html
        APP_KEY: "https://download01.logi.com/web/ftp/pub/techsupport/gaming/lghub_installer.exe"
    },
    "razer-synapse": {
        # From https://www.razer.com/synapse-3
        APP_KEY: "https://rzr.to/synapse-3-pc-download"
    },
    "nvidia-experience": {
        # From https://www.nvidia.com/en-us/geforce/geforce-experience/download/
        APP_KEY: "https://us.download.nvidia.com/GFE/GFEClient/3.27.0.120/GeForce_Experience_v3.27.0.120.exe"
    },
    "samsung-magician": {
        # From https://semiconductor.samsung.com/us/consumer-storage/support/tools/
        APP_KEY: "https://download.semiconductor.samsung.com/resources/software-resources/Samsung_Magician_installer_Official_8.0.0.900_Windows.zip"
    },
    "raspberry-pi-imager": {
        APP_KEY: "https://downloads.raspberrypi.org/imager/imager_latest.exe"
    },
    "notepad++": {
        # Update from https://github.com/notepad-plus-plus/notepad-plus-plus/releases
        APP_KEY: "https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v8.6/npp.8.6.Installer.x64.exe"
    },
    "firefox": {
        APP_KEY: "https://download.mozilla.org/?product=firefox-latest-ssl&os=win64&lang=en-US"
    },
    "chrome": {
        ARCHIVE_KEY: {
            PATHS_KEY: [os.path.join(os.getenv('USERPROFILE'), ".gitconfig")]
        },
        APP_KEY: "https://dl.google.com/chrome/install/ChromeStandaloneSetup64.exe"
    },
    "cygwin": {
        ARCHIVE_KEY: {
            PATHS_KEY: ["C:\\cygwin64\\home\\Jingleheimer\\.bash_history"]
        },
        APP_KEY: "https://www.cygwin.com/setup-x86_64.exe"
    },
    "virtualbox": {
        # Update from https://www.virtualbox.org/wiki/Downloads
        APP_KEY: "https://download.virtualbox.org/virtualbox/7.0.12/VirtualBox-7.0.12-159484-Win.exe"
    },
    "fusion360": {
        # Update from https://www.autodesk.com/products/fusion-360/appstream
        APP_KEY: "https://www.autodesk.com/products/fusion-360/appstream",
        SOURCE_KEY: "https://www.autodesk.com/products/fusion-360/appstream"
    },
    "veracrypt": {
        # Update from https://www.veracrypt.fr/en/Downloads.html
        APP_KEY: "https://launchpad.net/veracrypt/trunk/1.26.7/+download/VeraCrypt%20Setup%201.26.7.exe"
    },
    "teraterm": {
        # Update from https://github.com/TeraTermProject/teraterm/releases/latest
        APP_KEY: "https://github.com/TeraTermProject/teraterm/releases/download/v5.0/teraterm-5.0.exe"
    },
    "winscp": {
        # Update from https://winscp.net/eng/download.php
        APP_KEY: "https://winscp.net/download/WinSCP-6.1.2-Setup.exe"
    },
    "drive": {
        APP_KEY: "https://dl.google.com/drive-file-stream/GoogleDriveSetup.exe"
    },
    "gimp": {
        # Update from https://www.gimp.org/downloads/
        APP_KEY: "https://download.gimp.org/gimp/v2.10/windows/gimp-2.10.36-setup.exe"
    },
    "keepass": {
        APP_KEY: "https://sourceforge.net/projects/keepass/files/latest/download"
    },
    "steam": {
        APP_KEY: "https://cdn.akamai.steamstatic.com/client/installer/SteamSetup.exe"
    },
    "discord": {
        APP_KEY: "https://discord.com/api/downloads/distributions/app/installers/latest?channel=stable&platform=win&arch=x86"
    },
    "battle-net": {
        APP_KEY: "https://download.battle.net/?product=bnetdesk"
    },
    "epic": {
        APP_KEY: "https://launcher-public-service-prod06.ol.epicgames.com/launcher/api/installer/download/EpicGamesLauncherInstaller.msi"
    },

    # "C:\Users\Jingleheimer\Documents\My Games\DiRT Rally 2.0",
    # "C:\Users\Jingleheimer\Documents\Assetto Corsa",
    # "C:\Users\Jingleheimer\Documents\Project CARS 2",
    # "C:\Users\Jingleheimer\Documents\Adobe",
    # "C:\Users\Jingleheimer\Pictures\Lightroom"
}

# App specific keys
ENV_KEY = "environment"
ENV_PATH_KEY = "path"

VS_CODE_SETTING_PATH = os.path.join(os.getenv('APPDATA'), "Code", "User", "settings.json")
VS_CODE_APP_NAME = "code"
VS_CODE_KEY = "vscode"
VS_CODE_EXTENSIONS_KEY = "extensions"
VS_CODE_SETTINGS_KEY = "settings"

# Command line keys
BACKUP_OPERATION = "backup"
INSTALL_OPERATION = "install"
RESTORE_OPERATION = "restore"
OPERATIONS_LIST = [BACKUP_OPERATION, INSTALL_OPERATION, RESTORE_OPERATION]

###############################################################################################################################################################
# Logging
###############################################################################################################################################################
logging.basicConfig(format="[%(asctime)s.%(msecs)03d] [%(levelname)-8s] [%(module)s] --- %(message)s", datefmt="%Y-%m-%d %H:%M:%S")

class ShutdownHandler(logging.Handler):
    def emit(self, record):
            logging.shutdown()
            sys.exit(1)

logging.getLogger().addHandler(ShutdownHandler(level=logging.CRITICAL))


###############################################################################################################################################################
# Parse commandline arguments
###############################################################################################################################################################
parser = argparse.ArgumentParser(description='Tools for backing up and restoring important files before reimaging.',
                                 formatter_class = lambda prog: argparse.HelpFormatter(prog,max_help_position=40, width=120))
parser.add_argument('operation', choices=OPERATIONS_LIST, help='Operation to perform')
parser.add_argument('--archive', default='test.7z', help='Do a dry run of the operation')
parser.add_argument('-o', '--only', action='append', help='Process only this app for the operation')
parser.add_argument('-x', '--exclude', action='append', help='Applications to exclude from the operation.')
parser.add_argument('-e', '--encrypt', action='store_true', help='Encrypt the backup archive')
parser.add_argument('-d', '--dry-run', action='store_true', help='Do a dry run of the operation')
parser.add_argument('-v', '--verbose', action='store_true', help='Enable verbose output')

# Attempt to parse the arguments
try:
    args = parser.parse_args()
except argparse.ArgumentTypeError:
    parser.print_help()
    exit()

# Set the logging level based on input arguments
logging.getLogger().setLevel(logging.DEBUG if args.verbose else logging.INFO)

def CallCommand(command, shell, check=False, destructive=True):
    if destructive and args.dry_run:
        logging.info(f"Dry run: {command}")
        return ""
    else:
        logging.debug(f"Calling: {command}")
        result = subprocess.run(command, shell=shell, check=check, encoding='utf-8', stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        if result.returncode != 0:
            logging.error(f"{command[0]} returned an error: {result.returncode}\n{result.stdout}")

        logging.debug(f"Returned:\n{result.stdout}")
        return result.stdout

def WriteFile(path, data, dry_run_override=False):
    logging.info(f"Writing {path}")

    if args.dry_run and not dry_run_override:
        logging.debug(f"Dry run:\n{data}")
    else:
        with open(path, "w") as outfile:
            outfile.write(data)

def AddToArchive(path, excludes= []):
    arguments = [SEVEN_ZIP_PATH, "a", "-spf", args.archive, path] + excludes
    if args.encrypt :
        arguments.append(f"-p{password}")
    CallCommand(arguments, shell=False, destructive=False)

def RestoreFromArchive(files):
    arguments = [SEVEN_ZIP_PATH, "x", "-spf", args.archive, files]
    if args.encrypt :
        arguments.append(f"-p{password}")
    CallCommand(arguments, shell=False)

def TestArchive():
    password = ""

    info = CallCommand([SEVEN_ZIP_PATH, "l", args.archive, "-slt"], shell=False, destructive=False)
    if "Encrypted = +" in info:
        password = getpass("Archive appears encrypted, please provide password:")
    for attempt in range(5):
        try:
            result = CallCommand([SEVEN_ZIP_PATH, "t", args.archive, f"-p{password}"], shell=False, check=True, destructive=False)
            break
        except subprocess.CalledProcessError as e:
            if "Wrong password?" in e.stdout:
                password = getpass("Wrong password, try again:")
            else:
                logging.critical("Unknown error when attempting to check archive")

    return password

def BackupItem(metadata, info):
    filters = []
    if FILTERS_KEY in info:
        filters = info[FILTERS_KEY]

    for path in info[PATHS_KEY]:
        AddToArchive(path, filters)

def RestoreItem(metadata, info):
    for path in info[PATHS_KEY]:
        RestoreFromArchive(path)

def BackupVsCode(metadata):
    if not shutil.which(VS_CODE_APP_NAME):
        logging.error("VS Code not installed, skipping")
        return

    # Save installed Extensions
    extensions = CallCommand([VS_CODE_APP_NAME, "--list-extensions"], shell=True, destructive=False)
    metadata[VS_CODE_KEY] = {}
    metadata[VS_CODE_KEY][VS_CODE_EXTENSIONS_KEY] = extensions.split()

    # Save settings
    jsonFile = open(VS_CODE_SETTING_PATH, 'r') 
    settings = json.load(jsonFile)
    metadata[VS_CODE_KEY][VS_CODE_SETTINGS_KEY] = settings

def RestoreVsCode(metadata):
    if not shutil.which(VS_CODE_APP_NAME):
        logging.error("VS Code not installed, skipping")
        return

    # Reinstall extensions
    for extension in metadata[VS_CODE_KEY][VS_CODE_EXTENSIONS_KEY]:
        logging.info(f"Installing VS Code Extension: {extension}")
        CallCommand([VS_CODE_APP_NAME, "--install-extensions", extension], shell=True)
    
    # Restore settings
    WriteFile(VS_CODE_SETTING_PATH, json.dumps(metadata[VS_CODE_KEY][VS_CODE_SETTINGS_KEY], indent=4))

def BackupEnvironment(metadata):
    metadata[ENV_KEY] = {}
    
    key = OpenKey(HKEY_CURRENT_USER, r'Environment', 0, KEY_ALL_ACCESS)
    userPath = QueryValueEx(key, "Path")[0].split(';')
    metadata[ENV_KEY][ENV_PATH_KEY] = list(filter(lambda p: "S:" in p, userPath))

    # Since we know it, add 7z to the user path
    metadata[ENV_KEY][ENV_PATH_KEY].append(os.path.dirname(SEVEN_ZIP_PATH))

def RestoreEnvironment(metadata):
    user_path = ';'.join(metadata[ENV_KEY][ENV_PATH_KEY])
    if args.dry_run:
        logging.debug(f"Dry run: Set ENV Path to {user_path}")
    else:
        key = OpenKey(HKEY_CURRENT_USER, r'Environment', 0, KEY_ALL_ACCESS)
        SetValueEx(key, "Path", 0, REG_EXPAND_SZ, user_path)
        CloseKey(key)


###############################################################################################################################################################
# Main
###############################################################################################################################################################
if __name__=="__main__":
    metadata = {}
    password = ""
    backup_items = {}

    # Apply any filtering to backup items
    if args.exclude:
        backup_items = dict(filter(lambda a: a[0] not in args.exclude, BACKUP_ITEMS.items()))
    elif args.only:
        backup_items = dict(filter(lambda a: a[0] in args.only, BACKUP_ITEMS.items()))
    else:
        backup_items = BACKUP_ITEMS

    if args.operation == BACKUP_OPERATION:
        if args.encrypt:
            password = getpass()

        for name, info in backup_items.items():
            logging.info(f"Backing up {name}")
            if FUNCTIONS_KEY in info:
                locals()[info[FUNCTIONS_KEY][BACKUP_FUNCTION_KEY]](metadata)
            elif ARCHIVE_KEY in info:
                BackupItem(metadata, info[ARCHIVE_KEY])
            else:
                logging.debug(f"Skipping {name}, nothing to do.")

        # Store the metadata
        metadata = json.dumps(metadata, indent=4)
        logging.debug(f"metadata:\n{metadata}")
        WriteFile("metadata.json", metadata, dry_run_override=True)

    elif args.operation == INSTALL_OPERATION:

        # Load the metadata
        jsonFile = open('metadata.json', 'r') 
        metadata = json.load(jsonFile)

        # Test the archive and ensure it is OK, unencrypted, and ready.
        password = TestArchive()
        
        for name, info in backup_items.items():
            if APP_KEY in info:
                logging.info(f"Installing {name}")
                locals()[info[FUNCTIONS_KEY][RESTORE_FUNCTION_KEY]](metadata)
            else:
                logging.debug(f"Skipping {name}, nothing to do.")

    elif args.operation == RESTORE_OPERATION:

        # Load the metadata
        jsonFile = open('metadata.json', 'r') 
        metadata = json.load(jsonFile)

        # Test the archive and ensure it is OK, unencrypted, and ready.
        password = TestArchive()
        
        for name, info in backup_items.items():
            logging.info(f"Restoring {name}")
            if FUNCTIONS_KEY in info:
                locals()[info[FUNCTIONS_KEY][RESTORE_FUNCTION_KEY]](metadata)
            elif ARCHIVE_KEY in info:
                RestoreItem(metadata, info[ARCHIVE_KEY])
            else:
                logging.debug(f"Skipping {name}, nothing to do.")
    
    else:
        logging.critical(f"Unknown operation: {args.operation}")

    logging.info("Done")