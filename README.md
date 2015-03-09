Perma Word Plugin
=================

The Perma Word Plugin allows Perma.cc users to insert Perma Links directly from the right-click context menu of Microsoft Word.

## Installation

1. Download the "Perma Word Plugin.dotm" file from the Releases tab.
1. Open the file in Microsoft Word.
1. Enable macros.
1. Double-click to install.

## Compatibility

Believed working:

- Mac Word 2011
- Windows Word 2000

Known not working:

- Mac Word 2008 (no macro support)

Other versions are not yet tested -- let us know if you try.

## Building From Source

1. Clone this repository.
1. Open "Perma Word Plugin Base.docx" in Word.
1. Using "Save As", save a copy in the same directory as "Perma Word Plugin.dotm", using the "Word Macro-enabled Document (.dotm)" format.
1. Open the Visual Basic Editor from Word.
1. Import the "src/Build.bas" file into the "Perma Word Plugin.dotm" project.
1. Open Tools > References and enable "Microsoft Visual Basic for Application Extensibility 5.3".
1. Run the macro "ImportModules".
1. Rename the "Build1" module to "Build".
1. Save and close the template.

## Contributing

Changes should be committed only in the lib/ and src/ directories or "Perma Word Plugin Base.docx" file; do not commit .dotm files.

To export your changes to the lib/ and src/ directories, use the Build.ExportModules function.

To update your local .dotm with changes from upstream, first export your changes with ExportModules; then merge changes with git; then reimport using the ImportModules function.