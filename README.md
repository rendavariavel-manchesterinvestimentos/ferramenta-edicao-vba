# Vscode vba editing tool
Vscode tool with python and [xlwings library cli](https://docs.xlwings.org/en/latest/command_line.html) to easily edit and sync VBA code to excel. Based on this [xlwings tutorial](https://youtu.be/xoO-Fx0fTpM)

## How to use?
### 1. Download
Make sure [python](https://www.python.org/) is installed. Then [clone](https://git-scm.com/docs/git-clone) this repository to your machine using your preffered method.

### 2. Configure
Open the repo on [Vscode](https://code.visualstudio.com/) or [Cursor](https://www.cursor.com/). After that you can press `ctrl + shif + B` and select the `Configure Environment` vscode task, this is a one time action and you won't need to execute this anymore.

### 3. Usage (Vscode tasks)
To use this tool you need to be on Vscode or Cursor IDE's, the whole project it's based on Vscode tasks so it won't work on other IDE's.
By pressing `ctrl + shif + B` some options will appear:
1. Configure Environment
2. Create new project
3. Vba EDIT
4. Vba IMPORT

#### 3.1 Configure Environment
This task will configure the virtual environment (`.venv`) for the give python version you're using and install on this `.venv` [requirements](requirements.txt) needed for the project.

#### 3.2 Create a new project
This task will create on this repo a folder based on our [model](./docs/modelo) to concentrate all your VBA `.bas` files for a given `.xlsm`, and if you're using git by forking this repo, or another method, you have the advantage of version control.

#### 3.3 Vba EDIT
This task manages to open the `.xlsm` file and export all the existing VBA `.bas` on it, this is a destructive action so if you have other `.bas` files with the same name, they'll be subscripted. This task will prompt you for:
* What's the path where the VBA `.bas` will be stored.
* What's the path for the `.xlsm` file that will be edited.

#### 3.4 Vba IMPORT
This task manages to open the `.xlsm` file and import all the existing VBA `.bas` on a given directory. This task will prompt you for:
* What's the path where the VBA `.bas` are stored.
* What's the path for the `.xlsm` file that will be edited.

### 4 Troubleshooting
* xlwings errors: Verify if macros are enabled, see [this link](https://support.microsoft.com/en-us/office/enable-or-disable-macros-in-microsoft-365-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6)
* shell errors: Verify if python is installed, the `.venv` setted and [requirements](./requirements.txt) installed.
