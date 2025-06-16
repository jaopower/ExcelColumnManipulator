# ExcelColumnManipulator

ExcelColumnManipulator is a Python utility for advanced manipulation of Excel files using pandas. It provides a class and quick utility functions to rename, move, delete, merge, permute, and format columns in Excel spreadsheets.

## Features

- Rename columns
- Move columns to specific positions
- Delete columns by name, index, range, or pattern
- Merge columns (with separator, format, or custom logic)
- Permute (swap) columns based on conditions or patterns
- Change column values (single, multiple, or conditional)
- Create new columns (at index, with formulas, sequential, or from other columns)
- Format and standardize column values
- Format date columns

## Installation

Requires Python 3 and pandas.

```sh
pip install pandas openpyxl
```

## Usage

### As a Script

You can use the provided `computeData.py` script:

```sh
python computeData.py <excel_file>
```

### As a Library

Import and use the [`ExcelColumnManipulator`](ExcelColumnManipulator.py) class:

```python
from ExcelColumnManipulator import ExcelColumnManipulator

manipulator = ExcelColumnManipulator("your_file.xlsx")
manipulator.rename_columns({'OldName': 'NewName'})
manipulator.move_column('NewName', 0)
manipulator.save_excel("output.xlsx")
```

See the `main()` function in [`ExcelColumnManipulator.py`](ExcelColumnManipulator.py) for a full example.

## Quick Utility Functions

For common tasks, use the quick functions at the end of [`ExcelColumnManipulator.py`](ExcelColumnManipulator.py):

- `quick_rename_columns`
- `quick_move_column`
- `quick_reorder_columns`
- `quick_delete_column`
- `quick_delete_columns`
- `quick_merge_columns`
- `quick_permute_columns`
- `quick_change_values`
- `quick_create_column_at_index`
- `quick_create_sequential_column`
- `quick_format_date_columns`

## Example

```python
from ExcelColumnManipulator import quick_rename_columns

quick_rename_columns("input.xlsx", {'OldCol': 'NewCol'}, "output.xlsx")
```

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
