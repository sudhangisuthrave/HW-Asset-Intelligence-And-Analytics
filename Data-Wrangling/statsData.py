import pandas as pd


# Function to compute the statistics
def compute_statistics(file_path):
    # Read the CSV file
    df = pd.read_csv(file_path)

    # Check if required columns exist in the dataframe
    if 'id_acronym' not in df.columns or 'hostname' not in df.columns:
        raise ValueError("CSV file must contain 'id_acronym' and 'hostname' columns.")

    # Group by 'id_acronym' and count the number of 'hostnames' for each group
    result = df.groupby('id_acronym')['hostname'].count().reset_index()

    # Rename the columns for clarity
    pd.set_option('display.max_rows', 500)
    result.columns = ['id_acronym', 'hostnames_count']

    return result


# Main function to run the script
def main():
    # File path to the CSV file
    file_path = 'hostnames_06062024.csv'

    # Compute the statistics
    stats = compute_statistics(file_path)

    # Print the results
    print(stats)


# Run the main function
if __name__ == "__main__":
    main()

