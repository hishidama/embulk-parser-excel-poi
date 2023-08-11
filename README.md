# Apache POI Excel parser plugin for Embulk

Parses Microsoft Excel files(xls, xlsx) read by other file input plugins.  
This plugin uses Apache POI.

## Overview

* **Plugin type**: parser
* **Guess supported**: no
* Embulk 0.10 or later


## Example

```yaml
in:
  type: any file input plugin type
  parser:
    type: poi_excel
    sheets: ["DQ10-orb"]
    skip_header_lines: 1	# first row is header.
    columns:
    - {name: row, type: long, value: row_number}
    - {name: get_date, type: timestamp, cell_column: A, value: cell_value}
    - {name: orb_type, type: string}
    - {name: orb_name, type: string}
    - {name: orb_shape, type: long}
    - {name: drop_monster_name, type: string}
```

if omit **value**, specified `cell_value`.  
if omit **cell_column** when **value** is `cell_value`, specified next column.  


## Configuration

* **sheets**: sheet name. can use wildcards `*`, `?`. (list of string, required)
* **record_type**: record type.  (`row`, `column` or `sheet`. default: `row`)
* **skip_header_lines**: skip rows when **record_type**=`row` (skip columns when **record_type**=`column`). ignored when **record_type**=`sheet`. (integer, default: `0`)
* **columns**: column definition. see below. (hash, required)
* **sheet_options**: sheet option. see below. (hash, default: null)

### columns

* **name**: Embulk column name. (string, required)
* **type**: Embulk column type. (string, required)
* **value**: value type. see below. (string, default: `cell_value`)
* **column_number**: same as **cell_column**.
* **cell_column**: Excel column number. see below. (string, default: next column when **record_type**=`row`)
* **cell_row**: Excel row number. see below. (integer, default: next row when **record_type**=`column`)
* **cell_address**: Excel cell address such as `A1`, `Sheet1!B3`. (string, not required)
* **numeric_format**: format of numeric(double) to string such as `%4.2f`. (default: Java's Double.toString())
* **attribute_name**: use with value `cell_style`, `cell_font`, etc. see below. (list of string)
* **on_cell_error**: processing method of Cell error. see below. (string, default: `constant`)
* **formula_handling**: processing method of formula. see below. (`evaluate` or `cashed_value`. default: `evaluate`)
* **on_evaluate_error**: processing method of evaluate formula error. see below. (string, default: `exception`)
* **formula_replace**: replace formula before evaluate. see below.
* **on_convert_error**: processing method of convert error. see below. (string, default: `exception`)
* **search_merged_cell**: search merged cell when cell is BLANK. (`none`, `linear_search`, `tree_search` or `hash_search`, default: `hash_search`)

### value

* `cell_value`: value in cell.
* `cell_formula`: formula in cell. (if cell is not formula, same `cell_value`.)
* `cell_style`: all cell style attributes. returned json string. see **attribute_name**. (**type** required `string`)
* `cell_font`: all cell font attributes. returned json string. see **attribute_name**. (**type** required `string`)
* `cell_comment`: all cell comment attributes. returned json string. see **attribute_name**. (**type** required `string`)
* `cell_type`: cell type. returned Cell.getCellType() of POI.
* `cell_cached_type`: cell cached formula result type. returned Cell.getCachedFormulaResultType() of POI when CellType==FORMULA, otherwise same as `cell_type` (returned Cell.getCellType()).
* `file_name`: excel file name.
* `sheet_name`: sheet name.
* `row_number`: row number(1 origin).
* `column_number`: column number(1 origin).
* `constant`: constant value.

  * `constant.`*value*: specified value.
  * `constant`: null.

### cell_column

Basically used for **record_type**=`row`.

* `A`,`B`,`C`,...: column number of "A1 format".
* *number*: column number (1 origin).
* `+`: next column.
* `+`*name*: next column of name.
* `+`*number*: number next column.
* `-`: previous column.
* `-`*name*: previous column of name.
* `-`*number*: number previous column.
* `=`: same column.
* `=`*name*: same column of name.

### cell_row

Basically used for **record_type**=`column`.

* *number*: row number (1 origin).

### attribute_name

When **value** is `cell_style`, `cell_font`, or `cell_comment`, by default, it retrieves all attributes and converts them into a JSON string.
(Since it returns a JSON string, the **type** must be `string`.)

```yaml
    columns:
    - {name: foo, type: string, cell_column: A, value: cell_style}
```


By specifying the **attribute_name**, it retrieves only the specified attributes and converts them into a JSON string.

* **attribute_name**: attribute names. (list of string)

```yaml
    columns:
    - {name: foo, type: string, cell_column: A, value: cell_style, attribute_name: [border_top, border_bottom, border_left, border_right]}
```


Additionally, by appending a period after `cell_style` or `cell_font` and specifying the attribute name, you can retrieve only that attribute.  
In this case, it won't result in a JSON string, and you need to specify the type that matches the attribute's **type**.

```yaml
    columns:
    - {name: foo, type: long, value: cell_style.border}
    - {name: bar, type: long, value: cell_font.color}
```

In `cell_style` and `cell_font`, if **cell_column** is omitted, it targets the same column as the previous one.  
(In `cell_value`, omitting `cell_column` causes it to move to the next column.)


### on_cell_error

Processing method of Cell error (`#DIV/0!`, `#REF!`, etc).

```yaml
    columns:
    - {name: foo, type: string, cell_column: A, value: cell_value, on_cell_error: error_code}
```

* `constant`: set null. (default)
* `constant.`*value*: set specified value.
* `error_code`: set error code.
* `exception`: throw exception.


### formula_handling

Processing method of formula.

```yaml
    columns:
    - {name: foo, type: string, cell_column: A, value: cell_value, formula_handling: cashed_value}
```

* `evaluate`: evaluate formula. (default)
* `cashed_value`: cashed value in cell.


### on_evaluate_error

Processing method of evaluate formula error.

```yaml
    columns:
    - {name: foo, type: string, cell_column: A, value: cell_value, on_evaluate_error: constant}
```

* `constant`: set null.
* `constant.`*value*: set specified value.
* `exception`: throw exception. (default)


### formula_replace

Replace formula before evaluate.

```yaml
    columns:
    - {name: foo, type: string, cell_column: A, value: cell_value, formula_replace: [{regex: aaa, to: "A${row}"}, {regex: bbb, to: "B${row}"}]}
```

`${row}` is replaced with the current row number.
`${column}` is replaced with the current column string.


### on_convert_error

Processing method of convert error. ex) Excel boolean to Embulk timestamp

```yaml
    columns:
    - {name: foo, type: timestamp, format: "%Y/%m/%d", cell_column: A, value: cell_value, on_convert_error: constant.9999/12/31}
```

* `constant`: set null.
* `constant.`*value*: set specified value.
* `exception`: throw exception. (default)


### sheet_options

Options of individual sheet.

```yaml
  parser:
    type: poi_excel
    sheets: [Sheet1, Sheet2]
    columns:
    - {name: date, type: timestamp, cell_column: A}
    - {name: foo, type: string}
    - {name: bar, type: long}
    sheet_options:
      Sheet1:
        skip_header_lines: 1
        columns:
          foo: {cell_column: B}
          bar: {cell_column: C}
      Sheet2:
        skip_header_lines: 0
        columns:
          foo: {cell_column: D}
          bar: {value: constant.0}
```

**sheet_options** is map of sheet name.  
Map values are **skip_header_lines**, **columns**.

**columns** is map of column name.  
Map values are same **columns** in **parser** (excluding `name`, `type`).


## Install

1. download pom
   ```
   $ curl https://repo1.maven.org/maven2/io/github/hishidama/embulk/embulk-parser-excel-poi/0.2.0/embulk-parser-excel-poi-0.2.0.pom > embulk-parser-excel-poi-0.2.0.pom
   ```

2. install dependencies
   ```
   $ mvn install -f embulk-parser-excel-poi-0.2.0.pom
   ```

3. download and install jar
   ```
   $ export M2_REPO=$HOME/.m2/repository
   $ curl https://repo1.maven.org/maven2/io/github/hishidama/embulk/embulk-parser-excel-poi/0.2.0/embulk-parser-excel-poi-0.2.0.jar > $M2_REPO/io/github/hishidama/embulk/embulk-parser-excel-poi/0.2.0/embulk-parser-excel-poi-0.2.0.jar
   ```

4. add setting to $HOME/.embulk/embulk.properties
   ```
   plugins.parser.poi_excel=maven:io.github.hishidama.embulk:excel-poi:0.2.0
   ```


## Build

```
$ ./gradlew test
$ ./gradlew package
```

### Build to local Maven repository

```
./gradlew generatePomFileForMavenJavaPublication
mvn install -f build/publications/mavenJava/pom-default.xml
./gradlew publishToMavenLocal
```

