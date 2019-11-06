# excel2myfile

[![licence](https://img.shields.io/badge/Licence-MIT-green?style=flat-square)](https://github.com/crapStone/excel2myfile/blob/master/LICENSE) [![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg?style=flat-square)](https://github.com/psf/black)

A tool to extract data from excel files and write to complex via yaml config defined structures.

## Usage

```bash
usage: excel2myfile.py yaml_config excel_file out_file
```

## Yaml config

> use double `{{` or `}}` if you want a literal `{` or `}`.

```yaml
sheet_name: Sheet1  # workbook sheet
statement: "my statement with \n{sub1}"  # first statement wich is only inserted once, set to {} to use outest statement in every line
substatements:
  sub1: "\tmy statement with a nested substatement ({sub2}) and two placeholders {ph1}, {ph2}\n"
  sub2: "my nested substatement with a placeholder {ph3}"
placeholder_row_start: 2
placeholder_column_mapping:
  ph1: A
  ph2: B
  ph3: C
```
