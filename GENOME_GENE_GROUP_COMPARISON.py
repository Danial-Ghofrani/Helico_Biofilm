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


def find_universal_genes(df):
    return df.columns[(df.sum(axis=0) == df.shape[0])].tolist()


def create_group_gene_presence_table(df, groups, universal_genes):
    # Filter out universal genes
    df_filtered = df.drop(columns=universal_genes, errors='ignore')

    # Initialize results table
    result = {group: [] for group in groups}

    for group, genomes in groups.items():
        group_data = df_filtered.loc[df.index.isin(genomes)]
        for gene in df_filtered.columns:
            # Check if the gene is present in more than 50% of the group's genomes
            presence = group_data[gene].mean() > 0.5
            result[group].append(1 if presence else 0)

    # Create a DataFrame from the result
    result_df = pd.DataFrame(result, index=df_filtered.columns)

    # Save to Excel
    result_df.to_excel("group_gene_presence.xlsx")


def main():
    file_path = "non_universal_genes.xlsx"  # Change this to your actual file path

    df = load_data(file_path)
    groups = categorize_genomes(df)
    universal_genes = find_universal_genes(df)
    create_group_gene_presence_table(df, groups, universal_genes)


if __name__ == "__main__":
    main()
