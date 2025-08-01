import pandas as pd

# Load the original Excel file
file_path = 'data/surnames.xlsx'  # Replace with your file path
df = pd.read_excel(file_path, engine='openpyxl')

# Calculate the number of chunks
total_rows = len(df)
chunk_size = 1000
num_chunks = (total_rows + chunk_size - 1) // chunk_size

# Split the DataFrame into chunks and save each chunk as a new Excel file
for i in range(num_chunks):
    start_index = i * chunk_size
    end_index = (i + 1) * chunk_size
    chunk_df = df[start_index:end_index]
    
    # Generate the new file name (e.g., file_chunk_1.xlsx, file_chunk_2.xlsx, ...)
    output_file_path = file_path.replace('.xlsx', f'_{i + 1}.xlsx')
    
    chunk_df.to_excel(output_file_path, index=False, engine='openpyxl')

    print(f"Chunk {i + 1} saved to {output_file_path}")

print("Splitting complete!")