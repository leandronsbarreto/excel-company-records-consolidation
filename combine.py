import pandas as pd
import glob
import os


COLUMNS = [
    "Company Name",
    "ID Number",
    "Commercial Register Number",
    "Street and Number",
    "Postal Code",
    "City",
    "Phone",
    "Email Address"
]

OUTPUT_FILE = "consolidated.xlsx"


print("Current folder:")
print(os.getcwd())

files = glob.glob("*.xlsx")

# Ignore the final output file if it already exists.
files = [
    file for file in files
    if file.lower() != OUTPUT_FILE.lower()
]

print(f"Files found: {len(files)}")

for file in files:
    print(" -", file)

dataframes = []

for file in files:
    try:
        df = pd.read_excel(
            file,
            sheet_name=0,
            header=None,
            skiprows=1,
            engine="openpyxl"
        )

        df = df.dropna(how="all")

        print(
            f"{file}: {len(df.columns)} columns, "
            f"{len(df)} rows"
        )

        if len(df.columns) != len(COLUMNS):
            print(
                f"SKIPPED -> incorrect number of columns "
                f"({len(df.columns)} instead of {len(COLUMNS)})"
            )
            continue

        df.columns = COLUMNS
        dataframes.append(df)

    except Exception as error:
        print(f"ERROR in {file}: {error}")

print(f"\nValid DataFrames: {len(dataframes)}")

if not dataframes:
    raise Exception("No valid file was loaded.")

final_df = pd.concat(dataframes, ignore_index=True)

print(f"Total records: {len(final_df)}")

final_df.to_excel(
    OUTPUT_FILE,
    index=False,
    engine="openpyxl"
)

print(f"\nFile {OUTPUT_FILE} created successfully!")