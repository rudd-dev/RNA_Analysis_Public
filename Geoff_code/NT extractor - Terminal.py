import os
import sys
import xlsxwriter

# RNA nucleotide counter for FASTA files

def count_rna_nucleotides(fasta_file):
    """
    Reads a FASTA file and counts RNA nucleotides in sequences containing 'U'.

    Parameters:
        fasta_file (str): Path to the FASTA file.

    Returns:
        int: Total count of RNA nucleotides.
    """
    total_rna_nucleotides = 0

    with open(fasta_file, 'r') as file:
        sequence = ""
        for line in file:
            line = line.strip()
            if line.startswith(">"):
                # Process the current sequence
                if 'U' in sequence:
                    total_rna_nucleotides += len(sequence)
                sequence = ""  # Reset for the next sequence
            else:
                sequence += line.upper()

        # Handle the last sequence in the file
        if 'U' in sequence:
            total_rna_nucleotides += len(sequence)

    return total_rna_nucleotides

matrix = []
for file in range(1,len(sys.argv)):
    pdb = os.path.basename(sys.argv[file])[9:13]
    count = count_rna_nucleotides(sys.argv[file])
    matrix.append([pdb,count])

file_name = "Sequence_Length.xlsx"
desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
file_path = os.path.join(desktop_path, file_name)
workbook = xlsxwriter.Workbook(file_path)

output = workbook.add_worksheet()
output.write(0, 0, "PDB")
output.write(0, 1, "Length")

for i in range(0,len(matrix)):
    output.write_row(i+1, 0, matrix[i])

workbook.close()

print(f"Excel file '{file_name}' created on your desktop.")
print("Program Complete")