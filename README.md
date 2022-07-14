# Ecogy Plugins

This repository holds a bunch of internal Plugins/Commands for AutoCAD intended to be used internally by the
Ecogy team.

Currently, the best way to install is using the provided Installer, which can be found under the releases tab.
This will install the plugin into your Personal AutoCAD folder, and should be autoloaded on your next launch of AutoCAD.

## Uninstallation

You can uninstall the plugin in the normal Windows fashion: through the Apps and Featurs uninstall option. The program is under "Ecogy Plugins".

## Google Drive Plugin

This is an AutoCAD plugin that allows you to add a Google Drive path to your Places list.

### Usage

On initial setup (and any time you want to change the base path), you will need to run the `AddGoogleDrive` command
in the AutoCAD command line. This will ask you for the base path of the Google Drive, as well as for a depth.

This depth is how far back up the current path of your project you want to go to get a base path for the actual
place directory.

For example, assume your project is at C:\Users\User\AutoCAD\SolarPanel\Drawing.cfg, you can run `AddGoogleDrive`
and use the path G:\Google Drive\AutoCAD, and provide a depth of 1. This will set the place to G:\Google Drive\AutoCAD\SolarPanel,
and whenever you click on it it will drop you in that folder. This folder does not need to exist, it will be created for you.

Every project you open will change this path, but the base path and the depth will remain the same -- you can simply rerun
`AddGoogleDrive` if you would like to change the values.

## Spec Sheet Plugin

This is an AutoCAD plugin that allows you to import any PDF documents to use as Spec Sheets.

### Usage

Simply run the `ImportSpecSheet` command in the AutoCAD command line. A file selection dialog will appear, select all PDF
files you would like to import, and they will be added to the project.

## Filleting Plugin

This is an AutoCAD plugin that will fillet all Polylines in a drawing.

### Usage

Simply run the `FilletAll` command, supply the radius, and all of the polylines in your drawing will now be filleted.
