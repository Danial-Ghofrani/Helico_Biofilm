import pandas as pd

#### seperating common genes:
# Load the genome-gene table (first file)
genome_df = pd.read_excel("genome_gene.xlsx", header=0)  # Ensure the first row is used as column names

# Load the gene list from the third sheet of the second file
gene_list_df = pd.read_excel("gene_0_100_presence_results.xlsx", sheet_name=2, header=None)  # Read without headers

# Extract gene names, skipping the first row
selected_genes = gene_list_df.iloc[1:, 0].dropna().tolist()

# Ensure "genome_name" is included in the filtered columns
filtered_df = genome_df[['genome_name'] + [gene for gene in selected_genes if gene in genome_df.columns]]

# Count the number of present genes (sum of 1s in each row)
filtered_df["Total_Genes_Present"] = filtered_df[selected_genes].sum(axis=1)

# Create a DataFrame for the gene presence count
gene_presence_df = filtered_df[['genome_name', 'Total_Genes_Present']]

# Find the absent genes for each genome
absent_genes_dict = {
    row["genome_name"]: [gene for gene in selected_genes if row.get(gene, 1) == 0]
    for _, row in filtered_df.iterrows()
}

# Convert dictionary to DataFrame (for better Excel format)
absent_genes_df = pd.DataFrame(
    [(genome, ", ".join(genes)) for genome, genes in absent_genes_dict.items()],
    columns=["genome_name", "Absent_Genes"]
)

# Save all results in one Excel file with multiple sheets
output_filename = "filtered_genome_analysis_1.xlsx"
with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
    filtered_df.to_excel(writer, sheet_name="Filtered_Genes", index=False)
    gene_presence_df.to_excel(writer, sheet_name="Gene_Presence_Count", index=False)
    absent_genes_df.to_excel(writer, sheet_name="Absent_Genes", index=False)

print(f"Analysis saved in '{output_filename}'")