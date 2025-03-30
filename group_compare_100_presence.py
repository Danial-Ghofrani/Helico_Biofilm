import pandas as pd


def load_data(file_path):
    df = pd.read_excel(file_path)
    df.set_index(df.columns[0], inplace=True)  # Set genome names as index
    return df


def categorize_genomes(df):
    groups = {"high": [], "good": [], "moderate": [], "poor": []}
    for genome in df.index:
        if genome.startswith("H_"):
            groups["high"].append(genome)
        elif genome.startswith("G_"):
            groups["good"].append(genome)
        elif genome.startswith("M_"):
            groups["moderate"].append(genome)
        elif genome.startswith("P_"):
            groups["poor"].append(genome)
    return groups


def find_genes_present_in_all(df, groups):
    group_genes = {}

    for group, genomes in groups.items():
        group_data = df.loc[df.index.isin(genomes)]
        present_in_all = group_data.sum(axis=0) == len(genomes)  # Genes present in all genomes

        # Get the genes present in all genomes of the group
        genes_in_all = group_data.columns[present_in_all].tolist()

        group_genes[group] = genes_in_all

    return group_genes


def create_excel_file(group_genes, df, file_path):
    # Create a writer to save multiple sheets
    with pd.ExcelWriter(file_path) as writer:
        # First sheet: Genes present in all genomes for each group
        group_genes_df = pd.DataFrame.from_dict(group_genes, orient='index').transpose()
        group_genes_df.to_excel(writer, sheet_name="Genes_Present_in_All")

        # Second sheet: Matrix of gene presence for each group
        gene_list = sum(group_genes.values(), [])  # Flatten list of genes
        gene_list = list(set(gene_list))  # Remove duplicates

        # Initialize a DataFrame for gene presence matrix
        presence_matrix = {group: [] for group in group_genes}

        for group, genes in group_genes.items():
            group_data = df.loc[df.index.isin(group_genes[group])]
            for gene in gene_list:
                presence = 1 if gene in genes else 0
                presence_matrix[group].append(presence)

        # Create a DataFrame from the presence matrix
        presence_matrix_df = pd.DataFrame(presence_matrix, index=gene_list)
        presence_matrix_df.to_excel(writer, sheet_name="Gene_Presence_Matrix")


def main():
    file_path = "non_universal_genes.xlsx"  # Change this to your actual file path
    output_path = "group_nonuniversal_gene_analysis.xlsx"  # Output Excel file path

    df = load_data(file_path)
    groups = categorize_genomes(df)
    group_genes = find_genes_present_in_all(df, groups)
    create_excel_file(group_genes, df, output_path)
    print("Excel file has been created successfully.")


if __name__ == "__main__":
    main()

