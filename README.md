# Google Drive Plugin

This is an AutoCAD plugin that allows you to add a Google Drive path to your Places list.

## Usage

Currently, the best way to install is using the provided Installer (not currently in source control). This will install
the plugin into your Personal AutoCAD folder, and should be autoloaded on your next launch of AutoCAD.

On initial setup (and any time you want to change the base path), you will need to run the `AddGoogleDrive` command
in the AutoCAD command line. This will ask you for the base path of the Google Drive, as well as for a depth.

This depth is how far back up the current path of your project you want to go to get a base path for the actual
place directory.

For example, assume your project is at C:\Users\User\AutoCAD\SolarPanel\Drawing.cfg, you can run `AddGoogleDrive`
and use the path G:\Google Drive\AutoCAD, and provide a depth of 1. This will set the place to G:\Google Drive\AutoCAD\SolarPanel,
and whenever you click on it it will drop you in that folder. This folder does not need to exist, it will be created for you.

Every project you open will change this path, but the base path and the depth will remain the same -- you can simply rerun
`AddGoogleDrive` if you would like to change the values.
