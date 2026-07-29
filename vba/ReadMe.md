# VBA source

This folder contains exported VBA source files from NADABAS.

- `modules/` contains standard VBA modules (`.bas`)
- `classes/` contains class modules (`.cls`)
- `forms/` contains UserForms (`.frm`) and associated binary `.frx` files

The complete exports include component attributes and UserForm designer data.
VBA text files use CRLF line endings because Excel otherwise imports a
UserForm as an ordinary standard module. The repository's `.gitattributes`
enforces this on checkout.

The Excel add-in/workbook is not the primary source of truth. The exported VBA
files are versioned in Git.
