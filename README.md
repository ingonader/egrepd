# `egrepd`: `egrep` for documents

Extended regex grep tool for multiple document types. Filenames can be filtered using regex, too.



Invokation:

```bash
uvx egrepd -h
```



Supported document types:

* pdf
* odt
* odp
* docx
* doc (via `antiword` linux package)
* pptx (presentation and presenter notes)
* ppt (via `libreoffice` document converter)
* any other file type is treated as plain text file

Archive searching is optional. Use `-a` or `--include-archives` to search
archive contents after ordinary files. Supported formats include `.zip`,
`.jar`, `.7z`, `.tar`, `.tar.gz`, `.tgz`, `.tar.xz`, `.txz`, `.gz`, `.bz2`,
and `.xz`. Archive members use the same extension selection as regular files.
The `.7z` format is handled by the Python `py7zr` package.

```bash
uvx egrepd -a "search pattern" path/to/files
```

Has been developed using github copilot with human code review.
