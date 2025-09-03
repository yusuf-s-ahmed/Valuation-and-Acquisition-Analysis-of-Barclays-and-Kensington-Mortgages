import pandas as pd

# DCF template structure
inputs_data = {
    "Year": [2021, 2022, 2023, 2024],
    "Profit Before Tax": ["", "", "", ""],
    "Depreciation (PPE)": ["", "", "", ""],
    "Amortisation (Intangibles)": ["", "", "", ""],
    "One-off Adjustments": ["", "", "", ""],
    "Capital Expenditure": ["", "", "", ""],
    "Net Mortgage Cash Flows": ["", "", "", ""],
}

calculations_data = {
    "Year": [2021, 2022, 2023, 2024],
    "Profit Before Tax": ["=Inputs!B2", "=Inputs!B3", "=Inputs!B4", "=Inputs!B5"],
    "Depreciation": ["=Inputs!C2", "=Inputs!C3", "=Inputs!C4", "=Inputs!C5"],
    "Amortisation": ["=Inputs!D2", "=Inputs!D3", "=Inputs!D4", "=Inputs!D5"],
    "Adjustments": ["=Inputs!E2", "=Inputs!E3", "=Inputs!E4", "=Inputs!E5"],
    "Capex": ["=Inputs!F2", "=Inputs!F3", "=Inputs!F4", "=Inputs!F5"],
    "Mortgage CF": ["=Inputs!G2", "=Inputs!G3", "=Inputs!G4", "=Inputs!G5"],
    "Free Cash Flow": [
        "=SUM(B2:E2)-F2-G2",
        "=SUM(B3:E3)-F3-G3",
        "=SUM(B4:E4)-F4-G4",
        "=SUM(B5:E5)-F5-G5",
    ],
}

# Future projections (2025–2029)
projections_data = {
    "Year": [2025, 2026, 2027, 2028, 2029],
    "Projected FCF": ["", "", "", "", ""],
    "PV of FCF": ["", "", "", "", ""],
}

valuation_data = {
    "Item": [
        "PV of FCFs (2025–2029)",
        "PV of Terminal Value",
        "Enterprise Value",
        "Net Debt",
        "Equity Value",
        "Shares Outstanding",
        "Price per Share",
    ],
    "Value (£m)": ["", "", "", "", "", "", ""],
}

# Create DataFrames
df_inputs = pd.DataFrame(inputs_data)
df_calcs = pd.DataFrame(calculations_data)
df_proj = pd.DataFrame(projections_data)
df_val = pd.DataFrame(valuation_data)

# Save to Excel with multiple sheets
file_path = "Simple_DCF_Template.xlsx"
with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
    df_inputs.to_excel(writer, sheet_name="Inputs", index=False)
    df_calcs.to_excel(writer, sheet_name="Calculations", index=False)
    df_proj.to_excel(writer, sheet_name="Projections", index=False)
    df_val.to_excel(writer, sheet_name="Valuation", index=False)

file_path
