import pandas as pd

def load_data(path):
    return pd.read_excel(path)

def load_templates(path):
    df = pd.read_excel(path)
    templates = {}
    for _, row in df.iterrows():
        templates[row["ТК"]] = {
            "To": row["To"],
            "CC": row["CC"],
            "Subject": row["Subject"],
            "Body": row["Body"]
        }
    return templates

def filter_by_tc(data_df, tc_name):
    return data_df[data_df["ТК"] == tc_name]
