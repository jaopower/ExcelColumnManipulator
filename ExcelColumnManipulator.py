import pandas as pd
import os
from typing import Dict, List, Union

class ExcelColumnManipulator:
    def __init__(self, file_path: str, sheet_name: str = None):
        """
        Initialize the Excel manipulator
        
        Args:
            file_path (str): Path to the Excel file
            sheet_name (str): Name of the sheet to work with (default: first sheet)
        """
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.df = None
        self.load_excel()
    
    def load_excel(self):
        """Load Excel file into pandas DataFrame"""
        try:
            if self.sheet_name:
                self.df = pd.read_excel(self.file_path, sheet_name=self.sheet_name)
            else:
                self.df = pd.read_excel(self.file_path)
            print(f"Successfully loaded Excel file: {self.file_path}")
            print(f"Shape: {self.df.shape}")
            print(f"Columns: {list(self.df.columns)}")
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            raise
    
    def rename_columns(self, column_mapping: Dict[str, str]):
        """
        Rename columns in the DataFrame
        
        Args:
            column_mapping (dict): Dictionary mapping old column names to new names
                                 Example: {'old_name': 'new_name', 'col1': 'column_1'}
        """
        try:
            # Check if all old column names exist
            missing_cols = [col for col in column_mapping.keys() if col not in self.df.columns]
            if missing_cols:
                print(f"Warning: These columns don't exist: {missing_cols}")
            
            # Rename columns
            self.df.rename(columns=column_mapping, inplace=True)
            print(f"Renamed columns: {column_mapping}")
            print(f"New columns: {list(self.df.columns)}")
        except Exception as e:
            print(f"Error renaming columns: {e}")
            raise
    
    def move_column(self, column_name: str, position: Union[int, str]):
        """
        Move a column to a specific position
        
        Args:
            column_name (str): Name of the column to move
            position (int or str): Target position (index) or 'first'/'last'
        """
        try:
            if column_name not in self.df.columns:
                print(f"Error: Column '{column_name}' not found")
                return
            
            # Get the column data
            column_data = self.df[column_name]
            
            # Remove the column from its current position
            self.df.drop(columns=[column_name], inplace=True)
            
            # Insert at new position
            if position == 'first':
                self.df.insert(0, column_name, column_data)
            elif position == 'last':
                self.df[column_name] = column_data
            else:
                # Ensure position is within bounds
                position = max(0, min(position, len(self.df.columns)))
                self.df.insert(position, column_name, column_data)
            
            print(f"Moved column '{column_name}' to position {position}")
            print(f"New column order: {list(self.df.columns)}")
        except Exception as e:
            print(f"Error moving column: {e}")
            raise
    
    def move_multiple_columns(self, column_moves: List[tuple]):
        """
        Move multiple columns at once
        
        Args:
            column_moves (list): List of tuples (column_name, position)
                               Example: [('col1', 0), ('col2', 'last'), ('col3', 2)]
        """
        for column_name, position in column_moves:
            self.move_column(column_name, position)
    
    def reorder_columns(self, new_order: List[str]):
        """
        Reorder all columns according to a new order
        
        Args:
            new_order (list): List of column names in desired order
        """
        try:
            # Check if all columns are included
            missing_cols = [col for col in self.df.columns if col not in new_order]
            extra_cols = [col for col in new_order if col not in self.df.columns]
            
            if missing_cols:
                print(f"Warning: These existing columns not in new order: {missing_cols}")
                new_order.extend(missing_cols)
            
            if extra_cols:
                print(f"Warning: These columns don't exist: {extra_cols}")
                new_order = [col for col in new_order if col in self.df.columns]
            
            # Reorder columns
            self.df = self.df[new_order]
            print(f"Reordered columns: {list(self.df.columns)}")
        except Exception as e:
            print(f"Error reordering columns: {e}")
            raise
    
    def delete_column(self, column_name: str):
        """
        Delete a single column from the DataFrame
        
        Args:
            column_name (str): Name of the column to delete
        """
        try:
            if column_name not in self.df.columns:
                print(f"Error: Column '{column_name}' not found")
                return
            
            self.df.drop(columns=[column_name], inplace=True)
            print(f"Deleted column: '{column_name}'")
            print(f"Remaining columns: {list(self.df.columns)}")
        except Exception as e:
            print(f"Error deleting column: {e}")
            raise
    
    def delete_columns(self, column_names: List[str]):
        """
        Delete multiple columns from the DataFrame
        
        Args:
            column_names (list): List of column names to delete
        """
        try:
            # Check which columns exist
            existing_cols = [col for col in column_names if col in self.df.columns]
            missing_cols = [col for col in column_names if col not in self.df.columns]
            
            if missing_cols:
                print(f"Warning: These columns don't exist and will be skipped: {missing_cols}")
            
            if existing_cols:
                self.df.drop(columns=existing_cols, inplace=True)
                print(f"Deleted columns: {existing_cols}")
                print(f"Remaining columns: {list(self.df.columns)}")
            else:
                print("No valid columns to delete")
        except Exception as e:
            print(f"Error deleting columns: {e}")
            raise
    
    def delete_columns_by_pattern(self, pattern: str, case_sensitive: bool = False):
        """
        Delete columns that match a pattern (contains substring)
        
        Args:
            pattern (str): Pattern to match in column names
            case_sensitive (bool): Whether matching should be case sensitive
        """
        try:
            if case_sensitive:
                matching_cols = [col for col in self.df.columns if pattern in col]
            else:
                matching_cols = [col for col in self.df.columns if pattern.lower() in col.lower()]
            
            if matching_cols:
                self.df.drop(columns=matching_cols, inplace=True)
                print(f"Deleted columns matching pattern '{pattern}': {matching_cols}")
                print(f"Remaining columns: {list(self.df.columns)}")
            else:
                print(f"No columns found matching pattern '{pattern}'")
        except Exception as e:
            print(f"Error deleting columns by pattern: {e}")
            raise
    
    def delete_empty_columns(self):
        """
        Delete columns that are completely empty (all NaN or None values)
        """
        try:
            empty_cols = []
            for col in self.df.columns:
                if self.df[col].isna().all() or self.df[col].isnull().all():
                    empty_cols.append(col)
            
            if empty_cols:
                self.df.drop(columns=empty_cols, inplace=True)
                print(f"Deleted empty columns: {empty_cols}")
                print(f"Remaining columns: {list(self.df.columns)}")
            else:
                print("No empty columns found")
        except Exception as e:
            print(f"Error deleting empty columns: {e}")
            raise
    
    def delete_column_at_index(self, index: int):
        """
        Delete column at a specific index position
        
        Args:
            index (int): Index position of the column to delete (0-based)
        """
        try:
            if index < 0 or index >= len(self.df.columns):
                print(f"Error: Index {index} is out of range. Valid range: 0-{len(self.df.columns)-1}")
                return
            
            column_name = self.df.columns[index]
            self.df.drop(columns=[column_name], inplace=True)
            print(f"Deleted column at index {index}: '{column_name}'")
            print(f"Remaining columns: {list(self.df.columns)}")
        except Exception as e:
            print(f"Error deleting column at index: {e}")
            raise
    
    def delete_columns_at_indices(self, indices: List[int]):
        """
        Delete columns at multiple index positions
        
        Args:
            indices (list): List of index positions to delete (0-based)
        """
        try:
            # Sort indices in descending order to avoid index shifting issues
            valid_indices = []
            invalid_indices = []
            
            for idx in indices:
                if 0 <= idx < len(self.df.columns):
                    valid_indices.append(idx)
                else:
                    invalid_indices.append(idx)
            
            if invalid_indices:
                print(f"Warning: These indices are out of range and will be skipped: {invalid_indices}")
                print(f"Valid range: 0-{len(self.df.columns)-1}")
            
            if valid_indices:
                # Sort in descending order to delete from right to left
                valid_indices.sort(reverse=True)
                deleted_columns = []
                
                for idx in valid_indices:
                    column_name = self.df.columns[idx]
                    deleted_columns.append(f"'{column_name}' (index {idx})")
                    self.df.drop(columns=[column_name], inplace=True)
                
                print(f"Deleted columns: {deleted_columns}")
                print(f"Remaining columns: {list(self.df.columns)}")
            else:
                print("No valid indices to delete")
        except Exception as e:
            print(f"Error deleting columns at indices: {e}")
            raise
    
    def delete_column_range(self, start_index: int, end_index: int):
        """
        Delete a range of columns by index positions
        
        Args:
            start_index (int): Starting index (inclusive)
            end_index (int): Ending index (inclusive)
        """
        try:
            if start_index < 0 or end_index >= len(self.df.columns) or start_index > end_index:
                print(f"Error: Invalid range [{start_index}:{end_index}]. Valid range: 0-{len(self.df.columns)-1}")
                return
            
            # Get column names in the range
            columns_to_delete = []
            for i in range(start_index, end_index + 1):
                columns_to_delete.append(self.df.columns[i])
            
            self.df.drop(columns=columns_to_delete, inplace=True)
            print(f"Deleted columns from index {start_index} to {end_index}: {columns_to_delete}")
            print(f"Remaining columns: {list(self.df.columns)}")
        except Exception as e:
            print(f"Error deleting column range: {e}")
            raise
    
    def merge_columns(self, column_names: List[str], new_column_name: str, 
                     separator: str = " ", keep_original: bool = False,
                     handle_nan: str = "skip"):
        """
        Merge multiple columns into a single column
        
        Args:
            column_names (list): List of column names to merge
            new_column_name (str): Name for the new merged column
            separator (str): Separator to use between values (default: space)
            keep_original (bool): Whether to keep original columns (default: False)
            handle_nan (str): How to handle NaN values - 'skip', 'replace', or 'keep'
                            'skip': ignore NaN values
                            'replace': replace NaN with empty string
                            'keep': keep NaN as 'nan' string
        """
        try:
            # Check if all columns exist
            missing_cols = [col for col in column_names if col not in self.df.columns]
            if missing_cols:
                print(f"Error: These columns don't exist: {missing_cols}")
                return
            
            # Handle NaN values based on the specified method
            if handle_nan == "skip":
                # Skip NaN values during merge
                merged_series = self.df[column_names].apply(
                    lambda row: separator.join([str(val) for val in row if pd.notna(val)]), 
                    axis=1
                )
            elif handle_nan == "replace":
                # Replace NaN with empty string
                merged_series = self.df[column_names].fillna('').apply(
                    lambda row: separator.join([str(val) for val in row]), 
                    axis=1
                )
            else:  # handle_nan == "keep"
                # Keep NaN as 'nan' string
                merged_series = self.df[column_names].apply(
                    lambda row: separator.join([str(val) for val in row]), 
                    axis=1
                )
            
            # Add the new merged column
            self.df[new_column_name] = merged_series
            
            # Remove original columns if specified
            if not keep_original:
                self.df.drop(columns=column_names, inplace=True)
                print(f"Merged columns {column_names} into '{new_column_name}' and removed originals")
            else:
                print(f"Merged columns {column_names} into '{new_column_name}' and kept originals")
            
            print(f"Current columns: {list(self.df.columns)}")
        except Exception as e:
            print(f"Error merging columns: {e}")
            raise
    
    def merge_columns_with_format(self, column_names: List[str], new_column_name: str,
                                 format_string: str, keep_original: bool = False):
        """
        Merge columns using a custom format string
        
        Args:
            column_names (list): List of column names to merge
            new_column_name (str): Name for the new merged column
            format_string (str): Format string with placeholders like "{0} - {1}"
            keep_original (bool): Whether to keep original columns
        """
        try:
            # Check if all columns exist
            missing_cols = [col for col in column_names if col not in self.df.columns]
            if missing_cols:
                print(f"Error: These columns don't exist: {missing_cols}")
                return
            
            # Apply format string to each row
            def format_row(row):
                try:
                    values = [str(row[col]) if pd.notna(row[col]) else '' for col in column_names]
                    return format_string.format(*values)
                except:
                    return ""
            
            self.df[new_column_name] = self.df.apply(format_row, axis=1)
            
            # Remove original columns if specified
            if not keep_original:
                self.df.drop(columns=column_names, inplace=True)
                print(f"Merged columns {column_names} into '{new_column_name}' using format and removed originals")
            else:
                print(f"Merged columns {column_names} into '{new_column_name}' using format and kept originals")
            
            print(f"Current columns: {list(self.df.columns)}")
        except Exception as e:
            print(f"Error merging columns with format: {e}")
            raise
    
    def merge_columns_conditional(self, column_names: List[str], new_column_name: str,
                                 condition_func, keep_original: bool = False):
        """
        Merge columns with custom conditional logic
        
        Args:
            column_names (list): List of column names to merge
            new_column_name (str): Name for the new merged column
            condition_func: Custom function that takes a row and returns merged value
            keep_original (bool): Whether to keep original columns
        """
        try:
            # Check if all columns exist
            missing_cols = [col for col in column_names if col not in self.df.columns]
            if missing_cols:
                print(f"Error: These columns don't exist: {missing_cols}")
                return
            
            # Apply custom function to each row
            self.df[new_column_name] = self.df.apply(
                lambda row: condition_func(row[column_names]), axis=1
            )
            
            # Remove original columns if specified
            if not keep_original:
                self.df.drop(columns=column_names, inplace=True)
                print(f"Merged columns {column_names} into '{new_column_name}' with conditional logic and removed originals")
            else:
                print(f"Merged columns {column_names} into '{new_column_name}' with conditional logic and kept originals")
            
            print(f"Current columns: {list(self.df.columns)}")
        except Exception as e:
            print(f"Error merging columns conditionally: {e}")
            raise
    
    def concatenate_columns_numeric(self, column_names: List[str], new_column_name: str,
                                   operation: str = "sum", keep_original: bool = False):
        """
        Merge numeric columns using mathematical operations
        
        Args:
            column_names (list): List of numeric column names to merge
            new_column_name (str): Name for the new merged column
            operation (str): Operation to perform - 'sum', 'mean', 'max', 'min', 'product'
            keep_original (bool): Whether to keep original columns
        """
        try:
            # Check if all columns exist
            missing_cols = [col for col in column_names if col not in self.df.columns]
            if missing_cols:
                print(f"Error: These columns don't exist: {missing_cols}")
                return
            
            # Perform the specified operation
            if operation == "sum":
                self.df[new_column_name] = self.df[column_names].sum(axis=1)
            elif operation == "mean":
                self.df[new_column_name] = self.df[column_names].mean(axis=1)
            elif operation == "max":
                self.df[new_column_name] = self.df[column_names].max(axis=1)
            elif operation == "min":
                self.df[new_column_name] = self.df[column_names].min(axis=1)
            elif operation == "product":
                self.df[new_column_name] = self.df[column_names].prod(axis=1)
            else:
                print(f"Error: Unsupported operation '{operation}'. Use: sum, mean, max, min, product")
                return
            
            # Remove original columns if specified
            if not keep_original:
                self.df.drop(columns=column_names, inplace=True)
                print(f"Merged numeric columns {column_names} using '{operation}' into '{new_column_name}' and removed originals")
            else:
                print(f"Merged numeric columns {column_names} using '{operation}' into '{new_column_name}' and kept originals")
            
            print(f"Current columns: {list(self.df.columns)}")
        except Exception as e:
            print(f"Error merging numeric columns: {e}")
            raise
    
    def permute_columns_conditional(self, col1: str, col2: str, condition_col: str, 
                                   condition_value=None, condition_func=None, 
                                   comparison_operator: str = "=="):
        """
        Swap/permute values between two columns based on a condition from another column
        
        Args:
            col1 (str): First column name to swap
            col2 (str): Second column name to swap  
            condition_col (str): Column to check for condition
            condition_value: Value to compare against (used with comparison_operator)
            condition_func: Custom function that takes a value and returns True/False
            comparison_operator (str): Operator for comparison - "==", "!=", ">", "<", ">=", "<=", "in", "contains"
        """
        try:
            # Check if all columns exist
            required_cols = [col1, col2, condition_col]
            missing_cols = [col for col in required_cols if col not in self.df.columns]
            if missing_cols:
                print(f"Error: These columns don't exist: {missing_cols}")
                return
            
            # Create condition mask
            if condition_func is not None:
                # Use custom function
                condition_mask = self.df[condition_col].apply(condition_func)
                condition_desc = "custom function"
            elif condition_value is not None:
                # Use comparison operator
                if comparison_operator == "==":
                    condition_mask = self.df[condition_col] == condition_value
                elif comparison_operator == "!=":
                    condition_mask = self.df[condition_col] != condition_value
                elif comparison_operator == ">":
                    condition_mask = self.df[condition_col] > condition_value
                elif comparison_operator == "<":
                    condition_mask = self.df[condition_col] < condition_value
                elif comparison_operator == ">=":
                    condition_mask = self.df[condition_col] >= condition_value
                elif comparison_operator == "<=":
                    condition_mask = self.df[condition_col] <= condition_value
                elif comparison_operator == "in":
                    condition_mask = self.df[condition_col].isin(condition_value if isinstance(condition_value, list) else [condition_value])
                elif comparison_operator == "contains":
                    condition_mask = self.df[condition_col].astype(str).str.contains(str(condition_value), na=False)
                else:
                    print(f"Error: Unsupported comparison operator '{comparison_operator}'")
                    return
                condition_desc = f"{condition_col} {comparison_operator} {condition_value}"
            else:
                print("Error: Either condition_value or condition_func must be provided")
                return
            
            # Count rows that meet condition
            rows_to_swap = condition_mask.sum()
            
            if rows_to_swap == 0:
                print(f"No rows meet the condition: {condition_desc}")
                return
            
            # Perform the swap for rows that meet the condition
            temp_values = self.df.loc[condition_mask, col1].copy()
            self.df.loc[condition_mask, col1] = self.df.loc[condition_mask, col2]
            self.df.loc[condition_mask, col2] = temp_values
            
            print(f"Swapped values between '{col1}' and '{col2}' for {rows_to_swap} rows")
            print(f"Condition: {condition_desc}")
            
        except Exception as e:
            print(f"Error permuting columns: {e}")
            raise
    
    def permute_columns_multiple_conditions(self, col1: str, col2: str, conditions: List[dict]):
        """
        Swap values between two columns based on multiple conditions
        
        Args:
            col1 (str): First column name to swap
            col2 (str): Second column name to swap
            conditions (list): List of condition dictionaries, each containing:
                              {'column': str, 'value': any, 'operator': str}
        """
        try:
            # Check if all columns exist
            required_cols = [col1, col2] + [cond['column'] for cond in conditions]
            missing_cols = [col for col in required_cols if col not in self.df.columns]
            if missing_cols:
                print(f"Error: These columns don't exist: {missing_cols}")
                return
            
            # Create combined condition mask
            combined_mask = pd.Series([True] * len(self.df))
            
            for condition in conditions:
                col = condition['column']
                value = condition['value']
                operator = condition.get('operator', '==')
                
                if operator == "==":
                    mask = self.df[col] == value
                elif operator == "!=":
                    mask = self.df[col] != value
                elif operator == ">":
                    mask = self.df[col] > value
                elif operator == "<":
                    mask = self.df[col] < value
                elif operator == ">=":
                    mask = self.df[col] >= value
                elif operator == "<=":
                    mask = self.df[col] <= value
                elif operator == "in":
                    mask = self.df[col].isin(value if isinstance(value, list) else [value])
                elif operator == "contains":
                    mask = self.df[col].astype(str).str.contains(str(value), na=False)
                else:
                    print(f"Error: Unsupported operator '{operator}' in condition")
                    return
                
                combined_mask = combined_mask & mask
            
            rows_to_swap = combined_mask.sum()
            
            if rows_to_swap == 0:
                print("No rows meet all the specified conditions")
                return
            
            # Perform the swap
            temp_values = self.df.loc[combined_mask, col1].copy()
            self.df.loc[combined_mask, col1] = self.df.loc[combined_mask, col2]
            self.df.loc[combined_mask, col2] = temp_values
            
            condition_desc = " AND ".join([f"{c['column']} {c.get('operator', '==')} {c['value']}" for c in conditions])
            print(f"Swapped values between '{col1}' and '{col2}' for {rows_to_swap} rows")
            print(f"Conditions: {condition_desc}")
            
        except Exception as e:
            print(f"Error permuting columns with multiple conditions: {e}")
            raise
    
    def permute_columns_pattern_based(self, col1: str, col2: str, condition_col: str, 
                                     pattern: str, regex: bool = False):
        """
        Swap values between two columns based on pattern matching in another column
        
        Args:
            col1 (str): First column name to swap
            col2 (str): Second column name to swap
            condition_col (str): Column to check for pattern
            pattern (str): Pattern to match
            regex (bool): Whether pattern is a regular expression
        """
        try:
            # Check if all columns exist
            required_cols = [col1, col2, condition_col]
            missing_cols = [col for col in required_cols if col not in self.df.columns]
            if missing_cols:
                print(f"Error: These columns don't exist: {missing_cols}")
                return
            
            # Create condition mask based on pattern
            if regex:
                condition_mask = self.df[condition_col].astype(str).str.match(pattern, na=False)
                pattern_desc = f"regex pattern '{pattern}'"
            else:
                condition_mask = self.df[condition_col].astype(str).str.contains(pattern, na=False)
                pattern_desc = f"pattern '{pattern}'"
            
            rows_to_swap = condition_mask.sum()
            
            if rows_to_swap == 0:
                print(f"No rows match the {pattern_desc} in column '{condition_col}'")
                return
            
            # Perform the swap
            temp_values = self.df.loc[condition_mask, col1].copy()
            self.df.loc[condition_mask, col1] = self.df.loc[condition_mask, col2]
            self.df.loc[condition_mask, col2] = temp_values
            
            print(f"Swapped values between '{col1}' and '{col2}' for {rows_to_swap} rows")
            print(f"Pattern match: {pattern_desc} in column '{condition_col}'")
            
        except Exception as e:
            print(f"Error permuting columns with pattern: {e}")
            raise
    
    def change_column_values(self, column_name: str, old_value, new_value, 
                           comparison_operator: str = "==", case_sensitive: bool = True):
        """
        Change values in a column that match a specific condition
        
        Args:
            column_name (str): Name of the column to modify
            old_value: Value to search for and replace
            new_value: New value to replace with
            comparison_operator (str): How to match - "==", "!=", ">", "<", ">=", "<=", "contains", "startswith", "endswith"
            case_sensitive (bool): Whether string matching should be case sensitive
        """
        try:
            if column_name not in self.df.columns:
                print(f"Error: Column '{column_name}' not found")
                return
            
            # Create condition mask
            if comparison_operator == "==":
                if case_sensitive or not isinstance(old_value, str):
                    condition_mask = self.df[column_name] == old_value
                else:
                    condition_mask = self.df[column_name].astype(str).str.lower() == str(old_value).lower()
            elif comparison_operator == "!=":
                if case_sensitive or not isinstance(old_value, str):
                    condition_mask = self.df[column_name] != old_value
                else:
                    condition_mask = self.df[column_name].astype(str).str.lower() != str(old_value).lower()
            elif comparison_operator == ">":
                condition_mask = self.df[column_name] > old_value
            elif comparison_operator == "<":
                condition_mask = self.df[column_name] < old_value
            elif comparison_operator == ">=":
                condition_mask = self.df[column_name] >= old_value
            elif comparison_operator == "<=":
                condition_mask = self.df[column_name] <= old_value
            elif comparison_operator == "contains":
                if case_sensitive:
                    condition_mask = self.df[column_name].astype(str).str.contains(str(old_value), na=False)
                else:
                    condition_mask = self.df[column_name].astype(str).str.lower().str.contains(str(old_value).lower(), na=False)
            elif comparison_operator == "startswith":
                if case_sensitive:
                    condition_mask = self.df[column_name].astype(str).str.startswith(str(old_value), na=False)
                else:
                    condition_mask = self.df[column_name].astype(str).str.lower().str.startswith(str(old_value).lower(), na=False)
            elif comparison_operator == "endswith":
                if case_sensitive:
                    condition_mask = self.df[column_name].astype(str).str.endswith(str(old_value), na=False)
                else:
                    condition_mask = self.df[column_name].astype(str).str.lower().str.endswith(str(old_value).lower(), na=False)
            else:
                print(f"Error: Unsupported comparison operator '{comparison_operator}'")
                return
            
            rows_changed = condition_mask.sum()
            
            if rows_changed == 0:
                print(f"No values found matching condition: {column_name} {comparison_operator} {old_value}")
                return
            
            # Apply the change
            self.df.loc[condition_mask, column_name] = new_value
            
            print(f"Changed {rows_changed} values in column '{column_name}'")
            print(f"Condition: {column_name} {comparison_operator} {old_value} -> {new_value}")
            
        except Exception as e:
            print(f"Error changing column values: {e}")
            raise
    
    def change_values_multiple_replacements(self, column_name: str, replacement_dict: dict, 
                                          case_sensitive: bool = True):
        """
        Change multiple values in a column using a replacement dictionary
        
        Args:
            column_name (str): Name of the column to modify
            replacement_dict (dict): Dictionary mapping old values to new values
            case_sensitive (bool): Whether string matching should be case sensitive
        """
        try:
            if column_name not in self.df.columns:
                print(f"Error: Column '{column_name}' not found")
                return
            
            total_changes = 0
            change_summary = []
            
            for old_value, new_value in replacement_dict.items():
                if case_sensitive or not isinstance(old_value, str):
                    condition_mask = self.df[column_name] == old_value
                else:
                    condition_mask = self.df[column_name].astype(str).str.lower() == str(old_value).lower()
                
                rows_changed = condition_mask.sum()
                if rows_changed > 0:
                    self.df.loc[condition_mask, column_name] = new_value
                    total_changes += rows_changed
                    change_summary.append(f"{old_value} -> {new_value} ({rows_changed} rows)")
            
            if total_changes > 0:
                print(f"Made {total_changes} total changes in column '{column_name}':")
                for summary in change_summary:
                    print(f"  {summary}")
            else:
                print(f"No matching values found in column '{column_name}'")
                
        except Exception as e:
            print(f"Error changing multiple values: {e}")
            raise
    
    def change_values_conditional(self, target_column: str, new_value, 
                                condition_column: str, condition_value,
                                comparison_operator: str = "=="):
        """
        Change values in one column based on conditions in another column
        
        Args:
            target_column (str): Column to modify
            new_value: New value to set
            condition_column (str): Column to check condition against
            condition_value: Value to compare in condition column
            comparison_operator (str): How to compare - "==", "!=", ">", "<", ">=", "<=", "contains"
        """
        try:
            if target_column not in self.df.columns:
                print(f"Error: Target column '{target_column}' not found")
                return
            if condition_column not in self.df.columns:
                print(f"Error: Condition column '{condition_column}' not found")
                return
            
            # Create condition mask
            if comparison_operator == "==":
                condition_mask = self.df[condition_column] == condition_value
            elif comparison_operator == "!=":
                condition_mask = self.df[condition_column] != condition_value
            elif comparison_operator == ">":
                condition_mask = self.df[condition_column] > condition_value
            elif comparison_operator == "<":
                condition_mask = self.df[condition_column] < condition_value
            elif comparison_operator == ">=":
                condition_mask = self.df[condition_column] >= condition_value
            elif comparison_operator == "<=":
                condition_mask = self.df[condition_column] <= condition_value
            elif comparison_operator == "contains":
                condition_mask = self.df[condition_column].astype(str).str.contains(str(condition_value), na=False)
            else:
                print(f"Error: Unsupported comparison operator '{comparison_operator}'")
                return
            
            rows_changed = condition_mask.sum()
            
            if rows_changed == 0:
                print(f"No rows meet condition: {condition_column} {comparison_operator} {condition_value}")
                return
            
            # Apply the change
            self.df.loc[condition_mask, target_column] = new_value
            
            print(f"Changed {rows_changed} values in column '{target_column}' to '{new_value}'")
            print(f"Condition: {condition_column} {comparison_operator} {condition_value}")
            
        except Exception as e:
            print(f"Error changing values conditionally: {e}")
            raise
    
    def change_values_with_function(self, column_name: str, transform_function, 
                                   condition_function=None):
        """
        Change values in a column using custom functions
        
        Args:
            column_name (str): Name of the column to modify
            transform_function: Function to transform the value (takes old value, returns new value)
            condition_function: Optional function to determine which rows to change (takes value, returns True/False)
        """
        try:
            if column_name not in self.df.columns:
                print(f"Error: Column '{column_name}' not found")
                return
            
            if condition_function is None:
                # Apply to all rows
                self.df[column_name] = self.df[column_name].apply(transform_function)
                print(f"Applied transformation function to all values in column '{column_name}'")
            else:
                # Apply only to rows that meet condition
                condition_mask = self.df[column_name].apply(condition_function)
                rows_changed = condition_mask.sum()
                
                if rows_changed == 0:
                    print("No values meet the condition function")
                    return
                
                self.df.loc[condition_mask, column_name] = self.df.loc[condition_mask, column_name].apply(transform_function)
                print(f"Applied transformation function to {rows_changed} values in column '{column_name}'")
            
        except Exception as e:
            print(f"Error changing values with function: {e}")
            raise
    
    def standardize_values(self, column_name: str, standardization_type: str = "upper"):
        """
        Standardize values in a column (useful for data cleaning)
        
        Args:
            column_name (str): Name of the column to standardize
            standardization_type (str): Type of standardization - "upper", "lower", "title", "strip", "remove_spaces"
        """
        try:
            if column_name not in self.df.columns:
                print(f"Error: Column '{column_name}' not found")
                return
            
            original_values = self.df[column_name].copy()
            
            if standardization_type == "upper":
                self.df[column_name] = self.df[column_name].astype(str).str.upper()
            elif standardization_type == "lower":
                self.df[column_name] = self.df[column_name].astype(str).str.lower()
            elif standardization_type == "title":
                self.df[column_name] = self.df[column_name].astype(str).str.title()
            elif standardization_type == "strip":
                self.df[column_name] = self.df[column_name].astype(str).str.strip()
            elif standardization_type == "remove_spaces":
                self.df[column_name] = self.df[column_name].astype(str).str.replace(' ', '')
            else:
                print(f"Error: Unsupported standardization type '{standardization_type}'")
                return
            
            # Count changes
            changes = (original_values.astype(str) != self.df[column_name].astype(str)).sum()
            print(f"Standardized {changes} values in column '{column_name}' using '{standardization_type}'")
            
        except Exception as e:
            print(f"Error standardizing values: {e}")
            raise
    
    def create_column_at_index(self, index: int, column_name: str, default_value=None, 
                              values: List = None):
        """
        Create a new column at a specific index position
        
        Args:
            index (int): Index position where to insert the column (0-based)
            column_name (str): Name of the new column
            default_value: Default value for all rows (if values not provided)
            values (list): Specific values for the column (must match DataFrame length)
        """
        try:
            if index < 0 or index > len(self.df.columns):
                print(f"Error: Index {index} is out of range. Valid range: 0-{len(self.df.columns)}")
                return
                
            if column_name in self.df.columns:
                print(f"Error: Column '{column_name}' already exists")
                return
            
            # Determine column values
            if values is not None:
                if len(values) != len(self.df):
                    print(f"Error: Length of values ({len(values)}) doesn't match DataFrame length ({len(self.df)})")
                    return
                column_data = values
            else:
                column_data = [default_value] * len(self.df)
            
            # Insert the column at the specified index
            self.df.insert(index, column_name, column_data)
            
            print(f"Created column '{column_name}' at index {index}")
            print(f"Current columns: {list(self.df.columns)}")
            
        except Exception as e:
            print(f"Error creating column at index: {e}")
            raise
    
    def create_columns_at_indices(self, column_specs: List[dict]):
        """
        Create multiple columns at specific indices
        
        Args:
            column_specs (list): List of dictionaries with column specifications
                               Each dict should contain: {'index': int, 'name': str, 'default_value': any, 'values': list}
        """
        try:
            # Sort by index in descending order to avoid index shifting
            sorted_specs = sorted(column_specs, key=lambda x: x['index'], reverse=True)
            
            for spec in sorted_specs:
                index = spec['index']
                name = spec['name']
                default_value = spec.get('default_value', None)
                values = spec.get('values', None)
                
                if index < 0 or index > len(self.df.columns):
                    print(f"Warning: Skipping column '{name}' - index {index} out of range")
                    continue
                    
                if name in self.df.columns:
                    print(f"Warning: Skipping column '{name}' - already exists")
                    continue
                
                # Determine column values
                if values is not None:
                    if len(values) != len(self.df):
                        print(f"Warning: Skipping column '{name}' - values length mismatch")
                        continue
                    column_data = values
                else:
                    column_data = [default_value] * len(self.df)
                
                # Insert the column
                self.df.insert(index, name, column_data)
                print(f"Created column '{name}' at index {index}")
            
            print(f"Final columns: {list(self.df.columns)}")
            
        except Exception as e:
            print(f"Error creating multiple columns: {e}")
            raise
    
    def create_column_with_formula(self, index: int, column_name: str, formula_function):
        """
        Create a new column at index with values calculated using a formula function
        
        Args:
            index (int): Index position for the new column
            column_name (str): Name of the new column
            formula_function: Function that takes a row and returns the calculated value
        """
        try:
            if index < 0 or index > len(self.df.columns):
                print(f"Error: Index {index} is out of range. Valid range: 0-{len(self.df.columns)}")
                return
                
            if column_name in self.df.columns:
                print(f"Error: Column '{column_name}' already exists")
                return
            
            # Calculate values using the formula function
            calculated_values = self.df.apply(formula_function, axis=1)
            
            # Insert the column at the specified index
            self.df.insert(index, column_name, calculated_values)
            
            print(f"Created calculated column '{column_name}' at index {index}")
            print(f"Current columns: {list(self.df.columns)}")
            
        except Exception as e:
            print(f"Error creating column with formula: {e}")
            raise
    
    def create_column_from_other_columns(self, index: int, column_name: str, 
                                       source_columns: List[str], operation: str = "concat",
                                       separator: str = " "):
        """
        Create a new column at index based on other columns
        
        Args:
            index (int): Index position for the new column
            column_name (str): Name of the new column
            source_columns (list): List of source column names
            operation (str): Operation to perform - "concat", "sum", "mean", "max", "min"
            separator (str): Separator for concatenation (only used with "concat")
        """
        try:
            if index < 0 or index > len(self.df.columns):
                print(f"Error: Index {index} is out of range. Valid range: 0-{len(self.df.columns)}")
                return
                
            if column_name in self.df.columns:
                print(f"Error: Column '{column_name}' already exists")
                return
            
            # Check if source columns exist
            missing_cols = [col for col in source_columns if col not in self.df.columns]
            if missing_cols:
                print(f"Error: These source columns don't exist: {missing_cols}")
                return
            
            # Perform the specified operation
            if operation == "concat":
                # Concatenate with separator, skipping NaN values
                calculated_values = self.df[source_columns].apply(
                    lambda row: separator.join([str(val) for val in row if pd.notna(val)]), 
                    axis=1
                )
            elif operation == "sum":
                calculated_values = self.df[source_columns].sum(axis=1)
            elif operation == "mean":
                calculated_values = self.df[source_columns].mean(axis=1)
            elif operation == "max":
                calculated_values = self.df[source_columns].max(axis=1)
            elif operation == "min":
                calculated_values = self.df[source_columns].min(axis=1)
            else:
                print(f"Error: Unsupported operation '{operation}'. Use: concat, sum, mean, max, min")
                return
            
            # Insert the column at the specified index
            self.df.insert(index, column_name, calculated_values)
            
            print(f"Created column '{column_name}' at index {index} using '{operation}' on {source_columns}")
            print(f"Current columns: {list(self.df.columns)}")
            
        except Exception as e:
            print(f"Error creating column from other columns: {e}")
            raise
    
    def create_sequential_column(self, index: int, column_name: str, start_value: int = 1,
                               step: int = 1, prefix: str = "", suffix: str = ""):
        """
        Create a new column with sequential values at specific index
        
        Args:
            index (int): Index position for the new column
            column_name (str): Name of the new column
            start_value (int): Starting value for the sequence
            step (int): Step size for the sequence
            prefix (str): Prefix to add to each value
            suffix (str): Suffix to add to each value
        """
        try:
            if index < 0 or index > len(self.df.columns):
                print(f"Error: Index {index} is out of range. Valid range: 0-{len(self.df.columns)}")
                return
                
            if column_name in self.df.columns:
                print(f"Error: Column '{column_name}' already exists")
                return
            
            # Generate sequential values
            sequential_values = []
            for i in range(len(self.df)):
                value = start_value + (i * step)
                formatted_value = f"{prefix}{value}{suffix}"
                sequential_values.append(formatted_value)
            
            # Insert the column at the specified index
            self.df.insert(index, column_name, sequential_values)
            
            print(f"Created sequential column '{column_name}' at index {index}")
            print(f"Sequence: {start_value} to {start_value + (len(self.df)-1)*step} (step: {step})")
            print(f"Current columns: {list(self.df.columns)}")
            
        except Exception as e:
            print(f"Error creating sequential column: {e}")
            raise
    
    def format_date_columns(self, column_names: List[str] = None, date_format: str = "%m/%d/%Y"):
        """
        Format date columns to remove time component and standardize format
        
        Args:
            column_names (list): List of column names to format (if None, auto-detect date columns)
            date_format (str): Desired date format (default: "%m/%d/%Y" for MM/DD/YYYY)
        """
        try:
            if column_names is None:
                # Auto-detect datetime columns
                column_names = []
                for col in self.df.columns:
                    if pd.api.types.is_datetime64_any_dtype(self.df[col]):
                        column_names.append(col)
                
                if not column_names:
                    print("No datetime columns found for formatting")
                    return
                
                print(f"Auto-detected datetime columns: {column_names}")
            else:
                # Validate specified columns exist
                missing_cols = [col for col in column_names if col not in self.df.columns]
                if missing_cols:
                    print(f"Error: These columns don't exist: {missing_cols}")
                    return
            
            for col in column_names:
                try:
                    # Convert to datetime if not already
                    if not pd.api.types.is_datetime64_any_dtype(self.df[col]):
                        self.df[col] = pd.to_datetime(self.df[col], errors='coerce')
                    
                    # Format dates (remove time component)
                    self.df[col] = self.df[col].dt.strftime(date_format)
                    
                    print(f"Formatted column '{col}' to date format: {date_format}")
                    
                except Exception as e:
                    print(f"Warning: Could not format column '{col}': {e}")
            
            print(f"Date formatting completed for columns: {column_names}")
            
        except Exception as e:
            print(f"Error formatting date columns: {e}")
            raise
    
    def convert_to_date_only(self, column_names: List[str]):
        """
        Convert datetime columns to date-only format (removes time component)
        
        Args:
            column_names (list): List of column names to convert
        """
        try:
            missing_cols = [col for col in column_names if col not in self.df.columns]
            if missing_cols:
                print(f"Error: These columns don't exist: {missing_cols}")
                return
            
            for col in column_names:
                try:
                    # Convert to datetime if not already
                    if not pd.api.types.is_datetime64_any_dtype(self.df[col]):
                        self.df[col] = pd.to_datetime(self.df[col], errors='coerce')
                    
                    # Extract date component only
                    self.df[col] = self.df[col].dt.date
                    
                    print(f"Converted column '{col}' to date-only format")
                    
                except Exception as e:
                    print(f"Warning: Could not convert column '{col}': {e}")
            
        except Exception as e:
            print(f"Error converting to date-only: {e}")
            raise
    
    def display_info(self):
        """Display current DataFrame info"""
        print("\n" + "="*50)
        print("CURRENT DATAFRAME INFO")
        print("="*50)
        print(f"Shape: {self.df.shape}")
        print(f"Columns: {list(self.df.columns)}")
        print("\nFirst 5 rows:")
        print(self.df.head())
        print("="*50)
    
    def save_excel(self, output_path: str = None, sheet_name: str = 'Sheet1'):
        """
        Save the modified DataFrame to Excel
        
        Args:
            output_path (str): Path for output file (default: adds '_modified' to original name)
            sheet_name (str): Name of the sheet in output file
        """
        try:
            if output_path is None:
                # Create output filename by adding '_modified' before extension
                base_name = os.path.splitext(self.file_path)[0]
                extension = os.path.splitext(self.file_path)[1]
                output_path = f"{base_name}_modified{extension}"
            
            self.df.to_excel(output_path, sheet_name=sheet_name, index=False)
            print(f"File saved successfully: {output_path}")
        except Exception as e:
            print(f"Error saving file: {e}")
            raise

# Example usage and demonstration
def main():
    """Example usage of the ExcelColumnManipulator"""
    
    # Example 1: Basic usage
    print("EXAMPLE USAGE:")
    print("="*60)
    
    # Note: Replace 'your_file.xlsx' with actual file path
    file_path = "sample_data.xlsx"
    
    # Create sample data for demonstration
    sample_data = {
        'Name': ['John', 'Jane', 'Bob', 'Alice'],
        'Age': [25, 30, 35, 28],
        'City': ['NYC', 'LA', 'Chicago', 'Boston'],
        'Salary': [50000, 60000, 70000, 55000],
        'Department': ['IT', 'HR', 'Finance', 'Marketing']
    }
    
    # Create sample Excel file
    df_sample = pd.DataFrame(sample_data)
    df_sample.to_excel(file_path, index=False)
    print(f"Created sample file: {file_path}")
    
    # Initialize manipulator
    manipulator = ExcelColumnManipulator(file_path)
    manipulator.display_info()
    
    # Example operations
    print("\n1. Renaming columns...")
    manipulator.rename_columns({
        'Name': 'Employee_Name',
        'Age': 'Employee_Age',
        'Salary': 'Annual_Salary'
    })
    
    print("\n2. Moving 'Department' column to first position...")
    manipulator.move_column('Department', 'first')
    
    print("\n3. Moving 'Annual_Salary' to position 2...")
    manipulator.move_column('Annual_Salary', 2)
    
    print("\n4. Deleting 'City' column...")
    manipulator.delete_column('City')
    
    print("\n5. Deleting column at index 1...")
    manipulator.delete_column_at_index(1)
    
    print("\n6. Merging 'Employee_Name' and 'Department' columns...")
    manipulator.merge_columns(['Employee_Name', 'Department'], 'Name_Department', separator=' - ')
    
    # Add more sample data to demonstrate permutation
    print("\n7. Adding sample data with incorrect order for demonstration...")
    manipulator.df.loc[len(manipulator.df)] = ['Marketing', 25000, 'Alice Johnson - Marketing']
    manipulator.df.loc[len(manipulator.df)] = ['IT', 45000, 'Bob Smith - IT']
    
    # Let's say we want to swap the first two columns when Annual_Salary < 30000
    print("\n8. Swapping first two columns when Annual_Salary < 30000...")
    manipulator.permute_columns_conditional('Department', 'Annual_Salary', 'Annual_Salary', 30000, comparison_operator='<')
    
    print("\n9. Changing 'IT' department to 'Information Technology'...")
    manipulator.change_column_values('Department', 'IT', 'Information Technology')
    
    print("\n10. Standardizing department names to title case...")
    manipulator.standardize_values('Department', 'title')
    
    print("\n11. Creating a new ID column at index 0...")
    manipulator.create_sequential_column(0, 'Employee_ID', start_value=1001, prefix='EMP-')
    
    print("\n12. Creating a full name column at index 2...")
    def create_full_name(row):
        return f"{row['Name_Department'].split(' - ')[0]}"
    
    manipulator.create_column_with_formula(2, 'Full_Name', create_full_name)
    
    print("\n13. Final result:")
    manipulator.display_info()
    
    # Save the modified file
    manipulator.save_excel()
    
    print("\nDone! Check the '_modified' file for results.")

if __name__ == "__main__":
    main()

# Additional utility functions for quick operations
def quick_rename_columns(file_path: str, column_mapping: Dict[str, str], 
                        output_path: str = None):
    """Quick function to rename columns in an Excel file"""
    manipulator = ExcelColumnManipulator(file_path)
    manipulator.rename_columns(column_mapping)
    manipulator.save_excel(output_path)

def quick_move_column(file_path: str, column_name: str, position: Union[int, str], 
                     output_path: str = None):
    """Quick function to move a column in an Excel file"""
    manipulator = ExcelColumnManipulator(file_path)
    manipulator.move_column(column_name, position)
    manipulator.save_excel(output_path)

def quick_reorder_columns(file_path: str, new_order: List[str], 
                         output_path: str = None):
    """Quick function to reorder all columns in an Excel file"""
    manipulator = ExcelColumnManipulator(file_path)
    manipulator.reorder_columns(new_order)
    manipulator.save_excel(output_path)

def quick_delete_column(file_path: str, column_name: str, 
                       output_path: str = None):
    """Quick function to delete a column in an Excel file"""
    manipulator = ExcelColumnManipulator(file_path)
    manipulator.delete_column(column_name)
    manipulator.save_excel(output_path)

def quick_delete_columns(file_path: str, column_names: List[str], 
                        output_path: str = None):
    """Quick function to delete multiple columns in an Excel file"""
    manipulator = ExcelColumnManipulator(file_path)
    manipulator.delete_columns(column_names)
    manipulator.save_excel(output_path)

def quick_delete_column_at_index(file_path: str, index: int, 
                                output_path: str = None):
    """Quick function to delete a column at specific index in an Excel file"""
    manipulator = ExcelColumnManipulator(file_path)
    manipulator.delete_column_at_index(index)
    manipulator.save_excel(output_path)

def quick_delete_columns_at_indices(file_path: str, indices: List[int], 
                                   output_path: str = None):
    """Quick function to delete multiple columns at specific indices in an Excel file"""
    manipulator = ExcelColumnManipulator(file_path)
    manipulator.delete_columns_at_indices(indices)
    manipulator.save_excel(output_path)

def quick_merge_columns(file_path: str, column_names: List[str], new_column_name: str,
                       separator: str = " ", keep_original: bool = False,
                       output_path: str = None):
    """Quick function to merge columns in an Excel file"""
    manipulator = ExcelColumnManipulator(file_path)
    manipulator.merge_columns(column_names, new_column_name, separator, keep_original)
    manipulator.save_excel(output_path)

def quick_permute_columns(file_path: str, col1: str, col2: str, condition_col: str,
                         condition_value, comparison_operator: str = "==",
                         output_path: str = None):
    """Quick function to permute/swap columns based on condition in an Excel file"""
    manipulator = ExcelColumnManipulator(file_path)
    manipulator.permute_columns_conditional(col1, col2, condition_col, condition_value, 
                                           comparison_operator=comparison_operator)
    manipulator.save_excel(output_path)

def quick_change_values(file_path: str, column_name: str, old_value, new_value,
                       comparison_operator: str = "==", output_path: str = None):
    """Quick function to change column values in an Excel file"""
    manipulator = ExcelColumnManipulator(file_path)
    manipulator.change_column_values(column_name, old_value, new_value, comparison_operator)
    manipulator.save_excel(output_path)

def quick_create_column_at_index(file_path: str, index: int, column_name: str,
                                default_value=None, values: List = None,
                                output_path: str = None):
    """Quick function to create a column at specific index in an Excel file"""
    manipulator = ExcelColumnManipulator(file_path)
    manipulator.create_column_at_index(index, column_name, default_value, values)
    manipulator.save_excel(output_path)

def quick_create_sequential_column(file_path: str, index: int, column_name: str,
                                  start_value: int = 1, step: int = 1,
                                  prefix: str = "", suffix: str = "",
                                  output_path: str = None):
    """Quick function to create a sequential column at specific index in an Excel file"""
    manipulator = ExcelColumnManipulator(file_path)
    manipulator.create_sequential_column(index, column_name, start_value, step, prefix, suffix)
    manipulator.save_excel(output_path)

def quick_format_date_columns(file_path: str, column_names: List[str] = None,
                             date_format: str = "%m/%d/%Y", output_path: str = None):
    """Quick function to format date columns in an Excel file"""
    manipulator = ExcelColumnManipulator(file_path)
    manipulator.format_date_columns(column_names, date_format)
    manipulator.save_excel(output_path)