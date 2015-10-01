Perma Word Plugin
=================

The Perma Word Plugin allows Perma.cc users to insert Perma Links directly from the right-click context menu of Microsoft Word.

## Installation

1. Download the "Perma Word Plugin.dotm" file from the Releases tab.
2. Open the file in Microsoft Word.
3. Enable macros.
4. Double-click to install.

## Compatibility

Working:

- Windows Word 2013
- Windows Word 2010
- Mac Word 2011

Not working:

- Mac Word 2008 (no macro support)
- Mac Word 2016 (macro fails to modify context menu)

Other versions are not yet tested -- let us know if you try.

## Building From Source

1. Clone this repository.
2. Open "Perma Word Plugin Base.docx" in Word.
3. Using "Save As", save a copy in the same directory as "Perma Word Plugin.dotm", using the "Word Macro-enabled Document (.dotm)" format.
4. Open the Visual Basic Editor from Word.
5. Import the "src/Build.bas" file into the "Perma Word Plugin.dotm" project.
6. Open Tools > References and enable "Microsoft Visual Basic for Application Extensibility 5.3".
7. Run the macro "ImportModules".
8. Rename the "Build1" module to "Build".
9. Save and close the template.

## Contributing

Changes should be committed only in the lib/ and src/ directories or "Perma Word Plugin Base.docx" file; do not commit .dotm files.

To export your changes to the lib/ and src/ directories, use the Build.ExportModules function.

To update your local .dotm with changes from upstream, first export your changes with ExportModules; then merge changes with git; then reimport using the ImportModules function.
