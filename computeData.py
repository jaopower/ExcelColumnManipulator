# Basic usage
from ExcelColumnManipulator import ExcelColumnManipulator
import sys

if len(sys.argv) < 2:
    print("Usage: python computeData.py <excel_file>")
    sys.exit(1)

excel_file = sys.argv[1]
manipulator = ExcelColumnManipulator(excel_file)

# Rename columns
manipulator.rename_columns({'Libel': 'Libel de ligne', 'KRef': 'libel piece', 'Montant': 'Credit'})

# Delete column at index 3
manipulator.delete_column_at_index(3)

# Delete multiple columns
manipulator.delete_columns(['Sens', 'G', 'User', 'Pays'])

# Create a new column with the name 'Debit'
manipulator.create_column_at_index(6, 'Debit', default_value='')


# Swap when CpteCpta equals "411100"
manipulator.permute_columns_conditional(
    'Debit', 'Credit', 'CpteCpta', 411100, comparison_operator='contains'
)

# change column 411100 to 411 
manipulator.change_column_values('CpteCpta', 411100, 411)

# Basic merging with separator
manipulator.merge_columns(['CpteCpta', 'CpteTiers'], 'CpteCpta', separator='', keep_original=True)

# Delete a single column
manipulator.delete_column('CpteTiers')

# Move column to specific index
manipulator.move_column('CpteCpta', 2)
manipulator.move_column('libel piece', 5)

# Delete all empty columns
manipulator.delete_empty_columns()

# Replace VE by VT
manipulator.change_column_values('Jrnl', 'VE', 'VT')

# Format date columns 12/06/2025
manipulator.format_date_columns(['Date'], date_format="%d/%m/%Y")

# Save changes
manipulator.save_excel()