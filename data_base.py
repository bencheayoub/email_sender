import pandas as pd

df = pd.read_excel(r"omcList.xlsx")

# Combine first and last names into a single "name" column
df['name'] = df['first_name'].fillna('') + ' ' + df['last_name'].fillna('')

list_of_users = []

def users():
    for _, row in df.iterrows():
        # Skip row if any required field is missing
        if pd.isna(row["name"]) or pd.isna(row["email"]): # observation.
            continue

        list_of_users.append({
            "name": row["name"].strip(),   # Remove extra spaces
            "email": row["email"],
        })
    return list_of_users
