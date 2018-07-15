# GITVBA

*VBA program for checking in and out of your code to a source control repository

In programming, source control is an important, if not a time-saving / necessary job. In addition, source control brings all kinds of advantages in code review, version management and collaboration.

This is a familiar story for most programming languages such as Javascript, C++ etc.

Unfortunately, this is not the case for Visual Basic for Applications (abbreviated VBA).

Your VBA code sit in the VBA editor and is a time-consuming task to store all that code in separate files.

**GitVBA solves that for you.

You choose a desired repository location, program your VBA project. And if you are ready, check the project out to your repository folder.

Your Github repository registers all changes and your code are now in the source control flow.

And if you start with a new project, you can check in all objects. You can start immediately with your new project.

GitVBA exports the objects of your VBA project.
- All modules
- All classes
- All forms

GitVBA does not export the following objects:
- Worksheets
- Toolbar buttons

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See prerequisites and program use for notes on how to deploy the project.

### Prerequisites

This application needs the following references. Please check before running
* Microsoft Office 16.0 Object Library
* Microsoft Scripting Runtime
* Microsoft Visual Basic for Applications Extensibility 5.3

### Program use

- Start add-in GitVBA.xlam
- Choose your repository path
- Click on "Check In"-button to add all VBA codes to your repository in your active work book
- Click on "Check out"-button to check all your codes to active repository

## Built With

* Visual Basic for Applications 7.1 (Excel)

## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/EdgardeWit/GitVBA/tags). 

## Authors

* **Ron de Bruin** - *Initial work* - (https://www.rondebruin.nl/win/s9/win002.htm)
* See also the list of [contributors](https://github.com/EdgardeWit/GitVBA/contributors) who participated in this project.

## Acknowledgments

* If you also use your project as an add-in, use the "Create Add-in" function after Check out. The function "Check out" removes any existing add-in file.
* It is not possible to export toolbar buttons separately. These buttons are in your Excel file.
