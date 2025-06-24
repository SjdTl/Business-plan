import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import os as os
import numpy as np
from brokenaxes import brokenaxes
dir_path = os.path.dirname(os.path.realpath(__file__))
file_read = os.path.join(dir_path, f"total")
output_folder = os.path.join(dir_path, f"output")
# Load EXCEL
df = pd.read_excel(rf"{file_read}.xlsx", "total", header=0, engine="openpyxl", index_col=0)
df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
df = df.loc[:, ~df.columns.str.contains("Equation")]
# print(df)
def saveFig(name):
    sns.despine(left=True, bottom=True)
    plt.tight_layout()
    plt.savefig(os.path.join(output_folder, rf"{name}.svg"), transparent=True)

def RnD():
    # Filter out 'Total' if not needed
    colors = ["#d0e1f9", "#aec6cf", "#c2b280", "#cba135", "#b0b0b0"]

    df_RnD = df.loc[["3D printing", "Electronics", "Assembly labor", "Labor", "Other"]]
    df_melted = df_RnD.reset_index().melt(id_vars="index", var_name="Quarter", value_name="Amount")
    df_melted.rename(columns={"index": "Cost Category"}, inplace=True)
    # Extract year from quarter label
    df_melted["Year"] = df_melted["Quarter"].str.extract(r"(20\d{2})").astype(int)
    
    # Split dataframes
    df_2025 = df_melted[df_melted["Year"] == 2025]
    df_2026_plus = df_melted[df_melted["Year"] > 2025]

    # Aggregate 2026+ by year and cost category (average)
    df_yearly = df_2026_plus.groupby(["Cost Category", "Year"], as_index=False)["Amount"].sum()

    # Create 'Quarter' label for yearly data, e.g., '2026-Y'
    df_yearly["Quarter"] = df_yearly["Year"].astype(str)

    # Drop 'Year' (optional)
    df_yearly = df_yearly.drop(columns="Year")

    # Combine 2025 quarterly data and 2026+ yearly data
    df_combined = pd.concat([df_2025.drop(columns="Year"), df_yearly], ignore_index=True)

    plt.figure(figsize=(8, 4))
    sns.barplot(data=df_combined, x="Quarter", y=df_combined["Amount"]/1000 , hue="Cost Category", palette=colors)
    plt.ylabel("Cost (kEUR)")
    plt.xticks(rotation=45)
    saveFig("RnD")

def production_costs():
    # Melt for plotting
    df_prod = df.loc[["units", "costs", "cost/unit"]]
    df_melted = df_prod.reset_index().melt(id_vars="index", var_name="Quarter", value_name="Amount")
    df_pivot = df_melted.pivot(index="Quarter", columns="index", values="Amount").reset_index()
    # Plot: cost and units together
    fig, ax1 = plt.subplots(figsize=(8, 4))

    # Barplot for units (primary y-axis)
    sns.barplot(data=df_pivot, x="Quarter", y="units", color="lightblue", label="Units", ax=ax1)
    ax1.set_ylabel("Units", color="blue")
    ax1.tick_params(axis='x', rotation=45)

    # Secondary y-axis for costs and cost/unit
    ax2 = ax1.twinx()

    sns.lineplot(data=df_pivot, x="Quarter", y=df_pivot["costs"]/1000, color="darkred", marker="o", label="Costs (EUR)", ax=ax2)
    sns.lineplot(data=df_pivot, x="Quarter", y=df_pivot["cost/unit"]/1000, color="green", marker="o", label="Cost per Unit", ax=ax2)

    ax2.set_ylabel("Costs (kEUR)", color="darkred")

    # Legends (combine both axes)
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax2.legend(lines1 + lines2, labels1 + labels2, loc="upper left")
    saveFig("production_costs")

def service_costs():
    df_service = df.loc[["Marketing Cost", "Sales Cost", "Acquisition cost", "New Customers", "Customers Leaving", "Active Customers", "Units per Customer", "Cost per Acq. Customer"]]
    df_service.loc["Quarter"] = df_service.columns

    # df_service["Quarter"] = pd.Categorical(df_service["Quarter"], categories=quarters, ordered=True)
    fig, ax1 = plt.subplots(figsize=(8, 4))
    ax1.bar(df_service.loc["Quarter"] , (df_service.loc["Active Customers"]) -          (df_service.loc["New Customers"]), color="lightblue", label="Existing Customers")
    ax1.bar(df_service.loc["Quarter"] , (df_service.loc["New Customers"]), bottom=      (df_service.loc["Active Customers"]) - (df_service.loc["New Customers"]), color="dodgerblue", label="New Customers")
    ax1.bar(df_service.loc["Quarter"] , (df_service.loc["Customers Leaving"]), bottom=  (df_service.loc["Active Customers"]), color="red", label="Customers Leaving")
    # ax1.legend()
    ax2 = ax1.twinx()
    sns.lineplot(
        x=df_service.columns,
        y=df_service.loc["Marketing Cost"]/1000,
        color="darkred",
        marker="o",
        label="Marketing Cost",
        ax=ax2
    )
    sns.lineplot(
        x=df_service.columns,
        y=df_service.loc["Sales Cost"]/1000,
        color="green",
        marker="o",
        label="Sales Cost",
        ax=ax2
    )

    # Formatting
    ax1.set_ylabel("Amount", color="blue")
    ax2.set_ylabel("Costs (EUR)", color="darkred")
    ax1.tick_params(axis='x', rotation=45)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax2.legend(lines1 + lines2, labels1 + labels2, loc="upper left")

    saveFig("Service_costs")

def GnA():
    def quarter_to_halfyear(q):
        year, qtr = q.split('-')
        q_num = int(qtr[1])
        half = "H1" if q_num in [1, 2] else "H2"
        return f"{year}-{half}"
    
    categories = ["Office", "Finance", "Legal", "HR/Admin", "Software", "Insurance"]
    colors = ["#d0e1f9", "#aec6cf", "#c2b280", "#cba135", "#b0b0b0", "#a0c0e0"]

    df_gna = df.loc[categories]
    df_gna.loc["Quarter"] = df_gna.columns
    halfyears = [quarter_to_halfyear(q) for q in df_gna.loc["Quarter"]]
    df_halfyear = df_gna.T.groupby(by=halfyears).sum().T
    
    df_halfyear.loc["Halfyear"] = df_halfyear.columns
    fig, ax1 = plt.subplots(figsize=(8, 4))

    # Initialize stack base
    bottom = pd.Series([0] * df_halfyear.shape[1])

    # Plot stacked bars
    for cat, color in zip(categories, colors):
        # Divide by six since the values are given per month
        values = df_halfyear.loc[cat].astype(float).values /6
        ax1.bar(df_halfyear.loc["Halfyear"], values, bottom=bottom, label=cat, color=color)
        bottom += values

    # Finalize
    ax1.set_ylabel("Cost (EUR/MONTH)", color="blue")
    ax1.tick_params(axis='x', rotation=45)
    ax1.legend()

    saveFig("GnA")

def revenue():
    df_prod = df.loc[["units sold", "tutorials", "revenue", "Total received"]]
    df_melted = df_prod.reset_index().melt(id_vars="index", var_name="Quarter", value_name="Amount")
    df_pivot = df_melted.pivot(index="Quarter", columns="index", values="Amount").reset_index()
    # Plot: cost and units together
    fig, ax1 = plt.subplots(figsize=(8, 4))

    # Barplot for units (primary y-axis)
    sns.barplot(data=df_pivot, x="Quarter", y="units sold", color="#aec6cf", label="Total Units", ax=ax1)
    sns.barplot(data=df_pivot, x="Quarter", y=df_pivot["units sold"]*0.25, color="#31a8d3", label="Units with wireless extension", ax=ax1)
    sns.barplot(data=df_pivot, x="Quarter", y="tutorials", color="#cba135", label="Tutorials", ax=ax1)
    ax1.set_ylabel("Amount", color="blue")
    ax1.tick_params(axis='x', rotation=45)

    # Secondary y-axis for costs and cost/unit
    ax2 = ax1.twinx()

    sns.lineplot(data=df_pivot, x="Quarter", y=df_pivot["revenue"]/1000, color="darkred", marker="o", label="Revenue", ax=ax2)
    sns.lineplot(data=df_pivot, x="Quarter", y=df_pivot["Total received"]/1000, color="green", marker="o", label="Total received", ax=ax2)

    ax2.set_ylabel("Costs (kEUR)", color="darkred")

    # Legends (combine both axes)
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax2.legend(lines1 + lines2, labels1 + labels2, loc="upper left")

    saveFig("Revenue")

def costs():
    categories = ["R&D costs + maintenance", "Service", "Production", "Customer acquisition", "G&A"]

    colors = ["#d0e1f9", "#aec6cf", "#c2b280", "#cba135", "#b0b0b0"]

    df_costs = df.loc[categories]
    df_costs.loc["Quarter"] = df_costs.columns
    fig, ax1 = plt.subplots(figsize=(8, 4))

    # Initialize stack base
    bottom = pd.Series([0] * df_costs.shape[1])

    # Plot stacked bars
    for cat, color in zip(categories, colors):
        values = df_costs.loc[cat].astype(float).values / 1000
        ax1.bar(df_costs.loc["Quarter"], values, bottom=bottom, label=cat, color=color)
        bottom += values

    # Finalize
    ax1.set_ylabel("Cost (kEUR)", color="blue")
    ax1.tick_params(axis='x', rotation=45)
    ax1.legend()

    saveFig("costs")

def cashflow():
    df_cash = df.loc[["received", "cost", "cashflow", "cummulative cashflow"]]
    df_cash.loc["cost"] = -df_cash.loc["cost"]
    # Melt for plotting
    df_melted = df_cash.reset_index().melt(id_vars="index", var_name="Quarter", value_name="Amount")
    df_pivot = df_melted.pivot(index="Quarter", columns="index", values="Amount").reset_index()
    # Plot: cost and units together
    fig, ax1 = plt.subplots(figsize=(8, 4))
    

    # Barplot for units (primary y-axis)
    sns.barplot(data=df_pivot, x="Quarter", y=df_pivot["received"]/1000, color="lightgreen", label="Received", ax=ax1)
    sns.barplot(data=df_pivot, x="Quarter", y=df_pivot["cost"]/1000, color="#d35757", label="Cost", ax=ax1)
    ax1.set_ylabel("Net flow (EUR)", color="blue")
    ax1.tick_params(axis='x', rotation=45)

    sns.lineplot(data=df_pivot, x="Quarter", y=df_pivot["cashflow"]/1000, color="#4c8eea", marker="o", label="Cashflow", ax=ax1)
    sns.lineplot(data=df_pivot, x="Quarter", y=df_pivot["cummulative cashflow"]/1000, color="#284550", marker="o", label="Cumulative cashflow", ax=ax1)

    ax1.set_ylabel("Flow (kEUR)", color="blue")
    ax1.set_ylim(-100,300)
    ax1.legend()
    saveFig("cashflow")
    ax1.legend_.remove()
    fig.set_size_inches(8, 2)
    ax1.set_ylim(300,1200)
    saveFig("cashflow_high")

RnD()
production_costs()
service_costs()
GnA()
revenue()
costs()
cashflow()