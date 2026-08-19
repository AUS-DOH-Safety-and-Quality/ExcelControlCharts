# ExcelControlCharts
An Excel plugin for SPC charts and Funnel plots

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

### From Visual Studio Code

VS Code also provides good support for the plugin development workflow. Start by installing the [Office Add-ins Development Kit extension](https://marketplace.visualstudio.com/items?itemName=msoffice.microsoft-office-add-in-debugger).

Next, create a new `.vscode` folder in the root of the repo (if one does not already exist) and add the files (or append their contents to your own) in the [assets/vscode-configs](assets/vscode-configs) folder to it. This will provide a launch configuration for debugging the plugin in Excel.

You can launch the plugin with debugging support by pressing `F5` or selecting the "Preview Your Office Add-In" option from the Run menu:

<img width="628" height="281" alt="image" src="https://github.com/user-attachments/assets/28895eaf-f281-4cf1-9ed4-7d60deb6538b" />


This will perform the same steps as the `bun run start` command, but will also attach a debugger to the plugin - allowing for better support of logging and debugging:

<img width="1129" height="338" alt="image" src="https://github.com/user-attachments/assets/7b95172d-19d8-4710-b50e-b08e96974802" />

## Standalone Launcher

The add-in is hosted on GitHub Pages, so distributing it only requires sideloading the published manifest into Excel. `bun run start:deployed` does that from a checkout, and the same script compiles to a single executable for users who have Excel but no developer tooling:

```bash
bun run build:launcher       # release/ExcelControlCharts.exe (Windows, cross-compiles from any host)
bun run build:launcher:mac   # release/ExcelControlCharts (current platform)
```

The launcher downloads `manifest.xml` from the deployed site and registers it for sideloading, then exits. The add-in content itself is served by GitHub Pages.

To use it on a machine with Excel installed:

1. Copy `release/ExcelControlCharts.exe` to the target machine.
2. Run it once.
3. Open Excel and choose the add-in from the Home tab.

Set `ADDIN_BASE_URL` to point the launcher at a different deployment (a fork's Pages site, for example).

Important limitations:

- This is a desktop Excel sideloading flow, not AppSource publishing.
- The machine needs network access to the deployed site, both to build the launcher's manifest and to load the task pane.

