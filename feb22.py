import pandas as pd

def main():
    # Load the Excel file
    try:
        df = pd.read_excel('games_2022.xlsx', header=0)
    except Exception as e:
        print("Error reading the Excel file:", e)
        return

    # Print available column names for debugging
    print("Available Columns in Dataset:")
    print(df.columns.tolist())

    # Normalize column names (strip spaces, lowercase)
    df.columns = df.columns.str.strip().str.lower()

    # Expected columns
    required_columns = ['team', 'team_score', 'opponent_team_score']
    for col in required_columns:
        if col not in df.columns:
            print(f"Missing column: {col}")
            return

    # Find the D1 indicator column
    d1_column_name = None
    for col in df.columns:
        if "d1" in col or "division" in col or "tier" in col:
            d1_column_name = col
            break

    if not d1_column_name:
        print("No clear D1 indicator column found. Please manually check the dataset.")
        return

    # Convert D1 indicator column to boolean if necessary
    if df[d1_column_name].dtype == object:
        df['d1_bool'] = df[d1_column_name].str.lower().map({'yes': True, 'no': False})
    else:
        df['d1_bool'] = df[d1_column_name]

    # Determine win/loss status
    df['outcome'] = df.apply(lambda row: 'win' if row['team_score'] > row['opponent_team_score'] else 'loss', axis=1)

    # Filter out wins against D1 teams
    filtered_df = df[~((df['d1_bool'] == True) & (df['outcome'] == 'loss'))]

    # Group data by team
    team_stats = filtered_df.groupby('team').agg(
        total_points_scored=pd.NamedAgg(column='team_score', aggfunc='sum'),
        total_points_allowed=pd.NamedAgg(column='opponent_team_score', aggfunc='sum'),
        total_wins=pd.NamedAgg(column='outcome', aggfunc=lambda x: (x == 'win').sum()),
        total_losses=pd.NamedAgg(column='outcome', aggfunc=lambda x: (x == 'loss').sum())
    ).reset_index()

    # Print the results in the console
    print("\n=== Team Analysis Results ===")
    print(team_stats)

    # Save results to a CSV file
    team_stats.to_csv("team_analysis_results.csv", index=False)
    print("Results saved to team_analysis_results.csv. Download the file to view.")

if __name__ == '__main__':
    main()
