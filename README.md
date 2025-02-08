# HNTB

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

# To Activate Environment
- Run `build.sh` from terminal (the build.sh file has everything that is required to rebuild the environment or add new modules to the environment):
```
$ source build.sh
```
- A faster way to activate the environment is to use this script (once everything has been installed):
```
$ source ./venv/bin/activate
```
- Deactivate environment by running:
```
(venv) $ deactivate
```

# Examples

The following examples use the config yaml file
`./tests/artifacts/test_config.yaml`. This will pull dummy data from
`./tests/artifacts/hntb_dummy.xlsx` and save output to
`./tests/artifacts/Outputs/`. To use real data you will need to create a config
to pull data from the HNTB OneDrive. See `./config/courtney.yaml` or
`./config/courtney.yaml` for examples.

## Generate Run Face Sheets
- Activate the virtual environment (venv), then run:
```
(venv) $ python hntb_gen.py --config ./tests/artifacts/test_config.yaml --generate facesheets
```

## Generate PPT
- Activate the virtual environment (venv), then run:
```
(venv) $ python hntb_gen.py --config ./tests/artifacts/test_config.yaml --generate ppt
```

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
