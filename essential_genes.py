import pandas as pd

def load_data(file_path):
    # Load the Excel file containing the full results
    df = pd.read_excel(file_path, sheet_name=None)  # Read all sheets
    return df

def extract_relevant_columns(df, genes_to_extract):
    # Get the sheet names from the loaded Excel file
    sheet_name = list(df.keys())[0]  # Assuming the first sheet contains the relevant data

    # Extract the relevant sheet
    gene_presence_df = df[sheet_name]

    # Set the first column (genomes names) as the index
    gene_presence_df.set_index(gene_presence_df.columns[0], inplace=True)

    # Convert column names to lowercase
    gene_presence_df.columns = [col.lower() for col in gene_presence_df.columns]

    # Select only the columns of interest
    extracted_df = gene_presence_df[genes_to_extract]  # Select only the columns of interest

    return extracted_df

def save_to_excel(extracted_df, output_file_path):
    # Save the extracted data to a new Excel file
    extracted_df.to_excel(output_file_path)  # Saving with genome names as index (genomes in rows)

def main():
    input_file_path = "grouped_genome_gene.xlsx"  # Change this to your actual input file path
    output_file_path = "essential_genes_data.xlsx"  # Output file path

    genes_to_extract = [
        "cage", "flid", "lptb", "pgda", "flab", "tlpb", "luxs", "hspr", "rsfs"
    ]

    # Load the data from the previous Excel file
    df = load_data(input_file_path)

    # Extract the relevant columns (genes) from the first sheet
    extracted_df = extract_relevant_columns(df, genes_to_extract)

    # Save the extracted data to a new Excel file
    save_to_excel(extracted_df, output_file_path)
    print(f"Extracted gene data has been saved to {output_file_path}")

if __name__ == "__main__":
    main()
