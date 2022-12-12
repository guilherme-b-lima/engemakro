import abb_price
import siemens_price


manufacturer_dict = {
    1: {"name": "ABB", "function": abb_price.ABB_Price},
    2: {"name": "Siemens", "function": siemens_price.Siemens_Price},
    9: {"name": "Sair", "function": "Sair"},
}


def main():

    while True:

        for key in manufacturer_dict.keys():
            inner_dict1 = manufacturer_dict.get(key)
            menu = inner_dict1.get("name")
            print(key, "->", menu)

        user_input = input("Digite um dos códigos de fabricante acima: ")

        if user_input != "9":
            try:
                user_input = int(user_input)
                inner_dict2 = manufacturer_dict.get(user_input)
                inner_dict2.get("function")()
            except:
                print("ERRO: Código inválido\n")
        else:
            print("Saindo...")
            break


if __name__ == "__main__":
    main()
