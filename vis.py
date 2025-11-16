import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

plt.style.use("seaborn-v0_8")
sns.set_palette("muted")

BASE = "/Users/jeffff/Desktop/BPR/VCF-Visualization"

files = {
    "crime_type_14": f"{BASE}/cash-disbursements-by-crime-type-fy-14.xlsx",
    "crime_type_15": f"{BASE}/cash-disbursements-by-crime-type-fy-15-to-5.13.15.xlsx",
    "gender_14": f"{BASE}/cash-disbursements-by-gender-fy-14.xlsx",
    "gender_15": f"{BASE}/cash-disbursements-by-gender-fy-15-to-5.13.15.xlsx",
    "claims_14": f"{BASE}/claim-and-payment-stats-fy-14.xlsx",
    "claims_15": f"{BASE}/claim-and-payment-stats-fy-15-to-5.13.15.xlsx"
}

os.makedirs("figures", exist_ok=True)

def load_any_sheet(path):
    try:
        return pd.read_excel(path)
    except:
        xl = pd.ExcelFile(path)
        return xl.parse(xl.sheet_names[0])

df_crime_14 = load_any_sheet(files["crime_type_14"])
df_crime_15 = load_any_sheet(files["crime_type_15"])
df_gender_14 = load_any_sheet(files["gender_14"])
df_gender_15 = load_any_sheet(files["gender_15"])
df_claims_14 = load_any_sheet(files["claims_14"])
df_claims_15 = load_any_sheet(files["claims_15"])


def clean_cols(df):
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    return df

df_crime_14 = clean_cols(df_crime_14)
df_crime_15 = clean_cols(df_crime_15)
df_gender_14 = clean_cols(df_gender_14)
df_gender_15 = clean_cols(df_gender_15)
df_claims_14 = clean_cols(df_claims_14)
df_claims_15 = clean_cols(df_claims_15)


df_claims = pd.concat([df_claims_14.assign(year="2014"),
                       df_claims_15.assign(year="2015-")],
                      ignore_index=True)

def find_col(df, keywords):
    for col in df.columns:
        for k in keywords:
            if k in col:
                return col
    return None

col_total = find_col(df_claims, ["total", "claims"])
col_appr  = find_col(df_claims, ["approved"])
col_deny  = find_col(df_claims, ["denied"])

if col_appr and col_deny:
    df_claims["approval_rate"] = df_claims[col_appr] / df_claims[col_total]
    df_claims["denial_rate"] = df_claims[col_deny] / df_claims[col_total]

    plt.figure(figsize=(10,6))
    sns.lineplot(data=df_claims, x="year", y="approval_rate", marker="o", label="Approval Rate")
    sns.lineplot(data=df_claims, x="year", y="denial_rate", marker="o", label="Denial Rate")
    plt.ylabel("Rate")
    plt.title("Rhode Island Victim Compensation: Approval vs Denial Rate Over Time")
    plt.tight_layout()
    plt.savefig("figures/ri_approval_vs_denial.png", dpi=300)
    plt.close()

df_gender = pd.concat([df_gender_14.assign(year="2014"),
                       df_gender_15.assign(year="2015-")],
                      ignore_index=True)

col_gender = find_col(df_gender, ["gender"])
col_amt = find_col(df_gender, ["amount", "disbursements", "paid"])

if col_gender and col_amt:
    plt.figure(figsize=(8,6))
    sns.barplot(data=df_gender, x=col_gender, y=col_amt, estimator=sum)
    plt.title("Rhode Island Victim Compensation: Disbursements by Gender")
    plt.ylabel("Total $ Disbursed")
    plt.tight_layout()
    plt.savefig("figures/ri_gender_disbursements.png", dpi=300)
    plt.close()

df_crime = pd.concat([df_crime_14.assign(year="2014"),
                      df_crime_15.assign(year="2015-")],
                     ignore_index=True)

col_crime = find_col(df_crime, ["crime"])
col_amt2 = find_col(df_crime, ["amount", "disbursements"])

if col_crime and col_amt2:
    crime_sum = df_crime.groupby(col_crime)[col_amt2].sum().sort_values(ascending=False).head(12)

    plt.figure(figsize=(10,7))
    sns.barplot(x=crime_sum.values, y=crime_sum.index)
    plt.title("Rhode Island Victim Compensation: Top Crime Types by Disbursement")
    plt.xlabel("Total $ Disbursed")
    plt.tight_layout()
    plt.savefig("figures/ri_crime_type_disbursements.png", dpi=300)
    plt.close()
