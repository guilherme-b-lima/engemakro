import pandas as pd
from pathlib import Path
from tqdm import tqdm


# Selecting Siemens pricing list path
def load_source_df():
    # filepath_1 = input("Digite o caminho da lista de preços da ABB: "
    filepath_1 = "C:/Users/jamir/OneDrive/Área de Trabalho/Scripts/Manufacturer-price-update/spreadsheets/Siemens/2019_01_lista_PV.xlsx"
    # Creating origin dataframe from ABB pricing list, with headers starting at row 4.
    source_df = pd.read_excel(filepath_1, header=4)

    # Renaming the pricing and discount code columnn from source_df, to make it easier to work with
    source_df.rename(columns={"Preço líquido": "Price"}, inplace=True)
    source_df.rename(columns={"Grupo de material 2": "gdm2"}, inplace=True)
    return source_df


# Selecting agile software exported pricing list
def load_dest_df():
    # filepath_2 = input("Digite o caminho da lista a ser precificada: ")
    filepath_2 = "C:/Users/jamir/OneDrive/Área de Trabalho/Scripts/Manufacturer-price-update/spreadsheets/Siemens/SIEMENS-Banco dados 18-10-22 sem desconto.xlsx"
    # Creating destination dataframe from exported pricing list, with headers starting at row 0.
    dest_df = pd.read_excel(filepath_2, header=0, index_col="ID")
    return dest_df


def load_disc_df():
    # Selecting Siemens discount excel file
    # filepath_3 = input("Digite o caminho da lista de descontos ABB: ")
    filepath_3 = "C:/Users/jamir/OneDrive/Área de Trabalho/Scripts/Manufacturer-price-update/spreadsheets/Siemens/RelatorioDescontosCliente ENGEMAKRO 2016_2017.xlsx"
    disc_df = pd.read_excel(filepath_3, header=0)
    return disc_df


# Iterates over source_df, finding the component code in dest_df, then finding the discount code in disc_df,
# then applying the respective multiplier
def processing(source_df, dest_df, disc_df):
    dest_df_modified = dest_df
    count = 0
    for row in tqdm(
        source_df[["Material", "Price", "gdm2"]].itertuples(), total=len(source_df)
    ):
        found_1 = dest_df[dest_df["Codigo"] == getattr(row, "Material")]
        if not found_1.empty:
            price = row.Price
            count = count + 1
            if isinstance(price, str):
                price = 0
            found_2 = disc_df[disc_df["Grp_Mat_2"] == getattr(row, "gdm2")]
            if not found_2.empty:
                price *= found_2.iloc[0]["Multiplicador"]
            dest_df_modified.loc[found_1.index, "Preço"] = price

    return dest_df_modified, count


def output(dest_df_modified, count):
    Path("output").mkdir(parents=True, exist_ok=True)
    dest_df_modified.to_excel("output\Output_Siemens.xlsx")
    print(f"Finalizado, um total de {count} items foram encontrados\n")


def call_siemens():
    print("Cadastramento de preços da Siemens selecionado, aguarde...")
    source_df = load_source_df()
    dest_df = load_dest_df()
    disc_df = load_disc_df()
    [dest_df_modified, count] = processing(source_df, dest_df, disc_df)
    output(dest_df_modified, count)
