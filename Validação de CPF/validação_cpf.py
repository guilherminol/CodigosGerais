import pandas as pd

df = pd.read_excel("Entrada\CPF.xlsx")

def clean_cpf(cpf):
    cpf = str(cpf)
    for char in ",.- ":
        cpf = cpf.replace(char, "")
    return cpf

def is_valid_cpf(cpf):
    if len(cpf) != 11 or cpf == cpf[0] * 11:
        return False

    soma1, soma2 = 0, 0

    for i, digito in enumerate(cpf[:-2:]):
        soma1 += int(digito) * (i + 1)

    resto1 = soma1 % 11
    digito1 = 0 if resto1 == 10 else resto1

    for i, digito in enumerate(cpf[:-1]):
        soma2 += int(digito) * i


    resto2 = soma2 % 11
    digito2 = 0 if resto2 == 10 else resto2

    return cpf[-2:] == f"{digito1}{digito2}"


df["CPF VALIDADO"] = df["CPF - Corrigido"].apply(clean_cpf).apply(is_valid_cpf)
df["CPF VALIDADO"] = df["CPF VALIDADO"].map({True: "Válido", False: "Inválido"})


df.to_excel("Saida\CPF.xlsx", index=False)






