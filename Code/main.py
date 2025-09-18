from lectorCD import leer_creditos_debitos

def main():
    archivo = r"C:\Users\Mario\Downloads\excell debito y sistemas.xlsx"
    creditos, debitos = leer_creditos_debitos(archivo)

    for clave, valores in creditos.items():
        print("CREDITO:", clave, valores)

    for clave, valores in debitos.items():
        print("DEBITO:", clave, valores)

if __name__ == "__main__":
    main()