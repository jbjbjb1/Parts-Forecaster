# Scrambles the "Item Number" column so that it can no be recognised.

import pandas as pd
import random
import string

def create_cipher():
    """
    Creates a substitution cipher mapping.
    """
    letters = string.ascii_uppercase + string.digits
    shuffled_letters = list(letters)
    random.shuffle(shuffled_letters)
    cipher = dict(zip(letters, shuffled_letters))
    return cipher

def scramble_with_cipher(word, cipher):
    """
    Scrambles a word using the provided substitution cipher.
    """
    return ''.join(cipher.get(char, char) for char in word)

def scramble_part_numbers_with_cipher(part_numbers, cipher):
    """
    Scrambles a list of part numbers using a substitution cipher.
    Identical part numbers will result in the same scrambled output.
    """
    scrambled_dict = {}
    
    for part_number in part_numbers:
        if part_number not in scrambled_dict:
            scrambled_part_number = scramble_with_cipher(part_number, cipher)
            scrambled_dict[part_number] = scrambled_part_number
    
    return [scrambled_dict[part_number] for part_number in part_numbers]

def scramble_excel_column(input_file, output_file, column_name):
    """
    Reads an Excel file, scrambles the part numbers in the specified column using a substitution cipher,
    and saves the result to a new Excel file.
    """
    # Read the Excel file
    df = pd.read_excel(input_file)
    
    # Create a substitution cipher
    cipher = create_cipher()
    
    # Scramble the part numbers in the specified column
    df[column_name] = scramble_part_numbers_with_cipher(df[column_name].astype(str), cipher)
    
    # Save the result to a new Excel file
    df.to_excel(output_file, index=False)

# Example usage
input_file = "data\Dataset1.xlsx"  # Replace with your input file name
output_file = "data\scra_Dataset1.xlsx"  # Replace with your desired output file name
column_name = "Item Number"  # Replace with the actual column name in your Excel file

scramble_excel_column(input_file, output_file, column_name)

print(f"Scrambled part numbers have been saved to {output_file}")
