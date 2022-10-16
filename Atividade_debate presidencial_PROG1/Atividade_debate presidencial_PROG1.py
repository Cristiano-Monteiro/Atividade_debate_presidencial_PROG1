import pandas as pd
import csv

arquivo = pd.read_excel('debate_band.xlsx', usecols='F')

total_palavras = {}

for j in range(len(arquivo)):
    frase = arquivo['Texto'][j].replace('!', '').replace(',', '').replace('.', '').replace(':', '').replace('?', '').lower()
    frase_formatada = frase.split()
    for i in range(len(frase_formatada)):
        if(i == 0 or frase_formatada[i] not in total_palavras):
            total_palavras[frase_formatada[i]] = 0
        else:
            total_palavras[frase_formatada[i]] += 1

lista_palavras = []

for key in total_palavras.keys():
    if(total_palavras[key] != 0):
        lista = [total_palavras[key], key]
        lista_palavras.append(lista)

lista_palavras.sort(reverse=True)

lista_palavras_final = []

for l in lista_palavras:
    l.reverse()
    lista_palavras_final.append(l)

colunas = ['Palavra', 'Frequência']

linhas = lista_palavras_final

with open('frequencia_palavras.csv', 'w') as f:
    write = csv.writer(f)
    write.writerow(colunas)
    write.writerows(linhas)