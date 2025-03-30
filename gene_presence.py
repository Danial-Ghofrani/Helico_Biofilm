import pandas as pd

# Read the Excel file
file_path = r"/biofilm_former_Result.xlsx"
df = pd.read_excel(file_path, index_col=0)  # Set the first column as index

# Extract the percentage rows (rows with 'PP' in their labels)
df_percentages = df[df.index.str.endswith("PP")]

# Separate the first four PP rows (ignoring "Our genomes PP")
df_first_4_groups = df_percentages.iloc[:4]  # Selects the first 4 rows only

# Identify genes present in 100% of the genomes in each group
genes_present_100 = {}
genes_absent_0 = {}

for group in df_percentages.index:
    present_100 = df_percentages.loc[group][df_percentages.loc[group] == 100].index.tolist()
    absent_0 = df_percentages.loc[group][df_percentages.loc[group] == 0].index.tolist()

    genes_present_100[group] = present_100
    genes_absent_0[group] = absent_0

# Find genes that are 100% present in ALL FOUR groups (excluding 'Our genomes')
genes_100_in_all_4 = df_first_4_groups.columns[(df_first_4_groups == 100).all()].tolist()

# Convert results into DataFrames for exporting
df_present_100 = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in genes_present_100.items()]))
df_absent_0 = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in genes_absent_0.items()]))
df_genes_100_all_4 = pd.DataFrame({'Genes_100_in_All_4_Groups': genes_100_in_all_4})

# Define output file path
output_file = r"/gene_presence_results.xlsx"

# Save to an Excel file with three sheets
with pd.ExcelWriter(output_file) as writer:
    df_present_100.to_excel(writer, sheet_name="Genes_100_Present", index=False)
    df_absent_0.to_excel(writer, sheet_name="Genes_0_Absent", index=False)
    df_genes_100_all_4.to_excel(writer, sheet_name="Genes_100_in_All_4", index=False)

print(f"Results saved to {output_file}")
