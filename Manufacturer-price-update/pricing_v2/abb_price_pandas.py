import pandas as pd
from pathlib import Path
from tqdm import tqdm


# Selecting ABB pricing list pathh
def load_source_df():
    # filepath_1 = input("Digite o caminho da lista de preços da ABB: "
    filepath_1 = "./spreadsheets/ABB/Oficial Lista de Preços Setembro 22.xlsx"
    # Creating origin dataframe from ABB pricing list, with headers starting at row 2, and setting up "Item" column as the index.
    source_df = pd.read_excel(filepath_1, header=2, index_col="Item")

    # Renaming the pricing column from source_df, to make it easier to work with
    source_df.rename(columns={"Lista de  Preços": "Price"}, inplace=True)
    return source_df


# Selecting agile software exported pricing list
def load_dest_df():
    # filepath_2 = input("Digite o caminho da lista a ser precificada: ")
    filepath_2 = "./spreadsheets/ABB/ABB-Banco Dados 12-10-2022.xlsx"
    # Creating destination dataframe from exported pricing list, with headers starting at row 0.
    dest_df = pd.read_excel(filepath_2, header=0, index_col="ID")
    return dest_df


# Selecting ABB discount excel file
def load_disc_df():
    # filepath_3 = input("Digite o caminho da lista de descontos ABB: ")
    filepath_3 = "./spreadsheets/ABB/Desconto ABB.xlsx"
    disc_df = pd.read_excel(filepath_3, header=0)
    return disc_df


# Iterates over source_df, finding the component code in dest_df, then finding the GFD code (discount) in disc_df,
# then applying the respective multiplier
def processing(source_df, dest_df, disc_df):
    dest_df_modified = dest_df
    count = 0
    for row in tqdm(
        source_df[["Material", "Price", "gdm2"]].itertuples(), total=len(source_df)
    ):
        found_1 = dest_df[dest_df["Codigo"] == getattr(row, "Código")]
        if not found_1.empty:
            price = row.Price
            count = count + 1
            if isinstance(price, str):
                price = 0
            found_2 = disc_df[disc_df["Code"] == getattr(row, "GFD")]
            if not found_2.empty:
                price *= found_2.iloc[0]["Multiplier"]
            dest_df_modified.loc[found_1.index, "Preço"] = price

    return dest_df_modified, count


def output(dest_df_modified, count):
    Path("output").mkdir(parents=True, exist_ok=True)
    dest_df_modified.to_excel("output\Output_ABB.xlsx")
    print(f"Finalizado, um total de {count} items foram encontrados\n")


def call_abb():
    print("Cadastramento de preços da Abb selecionado, aguarde...")
    source_df = load_source_df()
    dest_df = load_dest_df()
    disc_df = load_disc_df()
    [dest_df_modified, count] = processing(source_df, dest_df, disc_df)
    output(dest_df_modified, count)
