import pandas as pd


def load_data(file_path):
    df = pd.read_excel(file_path)
    df.set_index(df.columns[0], inplace=True)  # Set genome names as index
    return df


def filter_genomes(df, min_genes=55, max_genes=62):
    gene_counts = df.sum(axis=1)  # Count number of present genes (1s) per genome
    filtered_genomes = df[(gene_counts >= min_genes) & (gene_counts <= max_genes)]  # Keep genomes with 55 to 62 genes
    return filtered_genomes


def find_missing_genes(filtered_df):
    missing_genes = filtered_df.loc[:, (filtered_df == 0).any(axis=0)]  # Keep only genes that have at least one 0
    return missing_genes


def save_to_excel(filtered_df, missing_genes, output_file="filtered_genomes.xlsx"):
    with pd.ExcelWriter(output_file) as writer:
        filtered_df.to_excel(writer, sheet_name="Filtered_Genomes")
        missing_genes.to_excel(writer, sheet_name="Missing_Genes")
    print(f"Filtered genomes and missing genes saved to {output_file}")


def main():
    input_file = "filtered_genome_table.xlsx"  # Update with actual file path

    df = load_data(input_file)

    # Remove genomes that have all 63 genes
    df_filtered = df[df.sum(axis=1) < 63]

    # Now filter genomes with 55 to 62 genes
    filtered_df = filter_genomes(df_filtered)
    missing_genes = find_missing_genes(filtered_df)

    save_to_excel(filtered_df, missing_genes)


if __name__ == "__main__":
    main()
