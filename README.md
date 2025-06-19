# HNTB
## [Video for how to use](https://youtube.com/playlist?list=PLafPl0mjmls3dsGBemAPvZQtYLLAkKkyz&si=M2xJqidNVczsFu3A)

The purpose of this repo is to simplify the use of EVMS Oto Head and Neck tumor
board scripts for multiple users. The original six scripts are located in the
directory `./original`. These scripts greatly accelerate the creation of
various documents used by HNTB. However, their current design makes it
challenging when multiple people use them. This is because:

1. Dependencies are not documented and easy to install. This is addressed by
the `build.sh` script, which creates a virtual environment and installs all
needed dependencies.
2. Each script requires its own set of config parameters, which are currently
hard coded in each file. This is addressed by the `hntb_gen.py` script, which
uses a single config file and is meant to replace the original six scripts.
**However, currently only `docs_FaceSheet.py` and `PPT.py` have been ported
over.**

## [Download Anaconda](https://www.anaconda.com/download)
- This is an easy way to make sure you have the correct version of python installed on your device
- Go to the Apple Icon at the top left of your screen and select About this Mac
- Determine if you have an Apple chip or an Intel chip
- Select the appropriate graphical installer from the Anaconda website
- Follow the installer instructions to complete the installation

## Download the code from this repository
- Save to your desired location on your computer

## Set up your folders
- Create a head and neck tumor board folder on your local computer
- Create a folder within the HNTB folder titled Outputs
- Unzip this repository and place the folder HNTB-main within the first folder your created (not in outputs)

## Create your config files
- Go to HNTB-main > config and open chris.yaml
```

# Input files
active_tumor_board_file: '/Users/chrishornung/Desktop/HNTB-Reboot/Active Tumor Board LINKED.xlsx'

# Output files
output_directory: '/Users/chrishornung/Desktop/HNTB-Reboot/Outputs'
ppt_filename: 'HNTB_PPT.pptx'
facesheets_filename: 'docs_facesheets.docx'
prelim_emails_prefix: 'PrelimEmails_'

# Template files
template_directory: './Templates'
header_image_filename: 'EVMSLogo.png'
facesheet_template_filename: 'FacesheetTemplate.docx'
ppt_template_filename: 'PPT_template.pptx'
```

- For #Input files, replace '/Users/chrishornung/Desktop/HNTB-Reboot/Active Tumor Board LINKED.xlsx' with your own file path to the first HNTB folder you created ex '/Users/JohnDoe/Desktop/YourFolderName/Active Tumor Board LINKED.xlsx'
- Make sure the path name is enclosed by ' '
- For #Output files change the path name to your username and HNTB folder ex '/Users/JohnDoe/Desktop/YourFolderName/Outputs'
-   The easiest way to find the path name is to copy paste from the bottom aspect of the Finder window. If it does not show up. Select View (at the top of your screen) > Show Path Bar. Once the path is visible at the bottom of the Finder window, you can copy and paste it into the .yaml file
- Save the file as YourName.yaml
- TROUBLESHOOTING - If it is not letting you save as a .yaml file you have two options.
-   1. Don't create a new text file, simply edit the file chris.yaml and save it. You can then rename the file or leave it as is, just make sure that the code to run the scripts below has the accurate config file name
    2. Click on the open text file then select Text edit at the top of your screen then select Settings and change the format to plain text. You will need to close the text editor and then repeat the process of copying and pasting the chris.yaml file into a new texteditor file and saving as your name .yaml. When you save, make sure that you uncheck the button that says "if no extension given, save at .txt"

## Download the Tumor board xlsx file
- You will ultimately use all of this code after you have edited the Active Tumor Board LINKED.xlsx document throughout the week, for the code to work, it needs this file
- Go to OneDrive then select File > Create a Copy > Download a Copy
- Once the file downloads, drag it into your original HNTB folder. Ensure that the name is Active Tumor Board LINKED.xlsx


## Set your directory
- Open Terminal then set your working directory with the below code where YourUserName and Location are where you saved the download for this repository
```
  cd /Users/YourUserName/Location/HNTB-main
```
ex. /Users/chrishornung/Desktop/HNTB-main

## Build environment 
- Run `build.sh` from terminal (the build.sh file has everything that is required to rebuild the environment or add new modules to the environment):
```
source build.sh
```
## Activate Environment
- Once everything has been installed, used the script below
```
source ./venv/bin/activate
```
- Deactivate environment by running:
```
deactivate
```

## Examples

The following examples use the config yaml file
`./tests/artifacts/test_config.yaml`. This will pull dummy data from
`./tests/artifacts/hntb_dummy.xlsx` and save output to
`./tests/artifacts/Outputs/`. To use real data you will need to use your config
to pull data from the HNTB OneDrive. See `./config/courtney.yaml` or
`./config/courtney.yaml` for examples.

### Generate Face Sheets
- Activate the virtual environment (venv) as above, then run:
Test
```
python hntb_gen.py --config ./tests/artifacts/test_config.yaml --generate facesheets
```
Real
```
python hntb_gen.py --config ./config/YourConfig.yaml --generate facesheets
```
- This file with save to the Outputs folder in your HNTB folder


### Generate PPT
- Activate the virtual environment (venv), then run:
Test
```
python hntb_gen.py --config ./tests/artifacts/test_config.yaml --generate ppt
```
Real
```
python hntb_gen.py --config ./config/YourConfig.yaml --generate ppt
````
- This file with save to the Outputs folder in your HNTB folder

### Generate Emails
- Activate the virtual environment (venv), then run:
Test
```
python hntb_gen.py --config ./tests/artifacts/test_config.yaml --generate emails
```
Real
```
python hntb_gen.py --config ./config/YourConfig.yaml --generate emails
````
- These files will save to the Outputs folder in your HNTB folder

# VSCode helpful actions:
- `command + /`: will turn code into comments

# Needed files, tables, ppt, etc..
- document with the facesheets to be handed out to the residents [docs_FaceSheet.py]
- power point presentation for the tumor board conference [HNTB_PPT.pptx]
- table with each attending's patients for the prelim emails [PrelimEmails.py]
- document with a list of all of the patients in our tumor board file for prelim emails
- document with a list of all of the patients, excluding pending patients for final emails
- document with a list of all of the radiology patients for final emails
- document with a list of all of the pathology patients for final emails

# Original Scripts
- docs_FaceSheet.py -> ./Outputs/docs_facesheets.docx
- PPT.py -> ./Outputs/HNTB_PPT.pptx
- PrelimEmails.py -> [Outputs/PrelimEmail_attending_name1.docx, ... ]
- FinalLists.py -> ?
- Rad_Path_FinalList.py -> ?
- Excel_FaceSheets.py -> ?
