import openpyxl
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# abrir a planilha
df_base = pd.read_excel("PastaAlunos.xlsx")

print(df_base)
#cada célula que contém as informações
for indice, linha in df_base.iterrows():
    nome_curso = linha["Nome do curso"]
    nome_participante = linha["Nome do participante"]
    tipo_participacao = linha["Tipo de participação"]
    carga_horaria = linha["Carga horária"]
    data_inicio = linha["Data início"]
    data_termino = linha["Data de término"]
    data_emissao = linha["Data de emissão do certificado"]

    print(nome_curso, nome_participante, tipo_participacao, data_inicio, data_termino)

    fonte_nome = ImageFont.truetype('./constanb.ttf',70)
    fonte_geral = ImageFont.truetype('./constan.ttf',50)
    fonte_data = ImageFont.truetype('./constan.ttf',40)

    image = Image.open('./Certificado.png')
    desenhar = ImageDraw.Draw(image)

    desenhar.text((670,666), nome_participante, fill='black',font=fonte_nome)
    desenhar.text((415,820), nome_curso, fill='black',font=fonte_geral)
    desenhar.text((534,907), tipo_participacao, fill='black',font=fonte_geral)
    desenhar.text((565,987),str(carga_horaria),fill='black',font=fonte_geral)

    data_inicio_str = data_inicio.strftime('%d/%m/%Y')
    data_termino_str = data_termino.strftime('%d/%m/%Y')
    data_emissao_str = data_emissao.strftime('%d/%m/%Y')

    desenhar.text((1555,823), data_inicio_str, fill='black', font=fonte_data)
    desenhar.text((1610,910), data_termino_str, fill='black', font=fonte_data)
    desenhar.text((1603,994), data_emissao_str, fill='black', font=fonte_data)

    image.save(f'./{indice} {nome_participante} certificado.png')