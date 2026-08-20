# ExcelControlCharts
An Excel plugin for SPC charts and Funnel plots


## Installing the Add-in

Installing the add-in requires downloading the add-in's metadata and registering it with Excel. We provide simple scripts to perform this registration for Excel Desktop, no admin rights needed.

The commands below will download the script and then execute it, you can also download the scripts from the [install](install) folder and run them yourself.


### Windows (PowerShell):

```pwsh
irm https://aus-doh-safety-and-quality.github.io/ExcelControlCharts/install.ps1 | iex
```

###  macOS (Terminal):

```sh
curl -fsSL https://aus-doh-safety-and-quality.github.io/ExcelControlCharts/install.sh | sh
```

Then restart Excel and choose the add-in from the Home tab.

## Uninstalling the Add-In

To remove, run the matching uninstall script:

### Windows (PowerShell):

```powershell
irm https://aus-doh-safety-and-quality.github.io/ExcelControlCharts/uninstall.ps1 | iex
```

####  macOS (Terminal):

```bash
curl -fsSL https://aus-doh-safety-and-quality.github.io/ExcelControlCharts/uninstall.sh | sh
```

## Using It Without Excel

The same charting panel is also published as a standalone web page, for when Excel isn't available:

**https://aus-doh-safety-and-quality.github.io/ExcelControlCharts/**

Type into cells, paste from Excel, or drop a CSV to load data — a sample dataset is included. Pick a
category and numerator column and the chart draws itself, updating live as you edit. Download it as
SVG or PNG. Everything runs in the browser — no data leaves your machine, and your workbook is saved
locally between visits.

The page also works offline: save `index.html` and open it directly, no server needed.

## Initialising the Development Environment

The repo uses submodules to include the [`PowerBI-SPC`](https://github.com/AUS-DOH-Safety-and-Quality/PowerBI-SPC) and [`PowerBI-Funnels`](https://github.com/AUS-DOH-Safety-and-Quality/PowerBI-Funnels) sources, so be sure to clone those when setting a local copy of the repo:

```bash
git clone --recursive https://github.com/AUS-DOH-Safety-and-Quality/ExcelControlCharts
```

The dependencies for the submodules are also included in the main `package.json` file, so you can install them all at once (note that this may take a few minutes):

```bash
cd ExcelControlCharts
bun install
```

## Developing Locally

### From The Command Line

To run the development server, use:

```bash
bun run start
```

This will compile the plugin and start a local server that you can use to test the plugin in Excel. A blank spreadsheet will open with the plugin loaded, but it will also be available in any existing spreadsheets you have open.

The same server also hosts the Excel-free web page at its root (`https://localhost:3100/`), so both
front ends rebuild together as you edit.

### From Visual Studio Code

VS Code also provides good support for the plugin development workflow. Start by installing the [Office Add-ins Development Kit extension](https://marketplace.visualstudio.com/items?itemName=msoffice.microsoft-office-add-in-debugger).

Next, create a new `.vscode` folder in the root of the repo (if one does not already exist) and add the files (or append their contents to your own) in the [assets/vscode-configs](assets/vscode-configs) folder to it. This will provide a launch configuration for debugging the plugin in Excel.

You can launch the plugin with debugging support by pressing `F5` or selecting the "Preview Your Office Add-In" option from the Run menu:

<img width="628" height="281" alt="image" src="https://github.com/user-attachments/assets/28895eaf-f281-4cf1-9ed4-7d60deb6538b" />


This will perform the same steps as the `bun run start` command, but will also attach a debugger to the plugin - allowing for better support of logging and debugging:

<img width="1129" height="338" alt="image" src="https://github.com/user-attachments/assets/7b95172d-19d8-4710-b50e-b08e96974802" />
