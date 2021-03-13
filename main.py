"""
Extrai algumas informações sobre o IPCA do site do IBGE e as dispõe em forma de uma planilha xls
as informações são:
    IPCA no mês corrente, IPCA acumulado nos últimos 12 meses, IPCA acumulado no ano;
        - os dados são decompostos em setores
    Histórico do IPCA desde 94;
    Histórico de pesos dos setores na composição do IPCA;
"""
import zipfile
import io
import requests
import openpyxl
import pandas as pd


URL = "https://sidra.ibge.gov.br/geratabela?format=xlsx&name=tabela7060.xlsx&terr=N&rank=-&query=t/7060/n1/all/v/all/p/all/c315/7169,7170,7445,7486,7558,7625,7660,7712,7766,7786/d/v63%202,v66%204,v69%202,v2265%202/l/,v%2Bt%2Bp,c315"

r = requests.get(URL)
r = r.content

raw = pd.ExcelFile(r)
raw = pd.read_excel(raw)
#raw.head

categoria = raw.iloc[4:14, 0]
categoria = categoria.rename()
categoria = categoria.reset_index(drop=True)
categoria_str = []
for i in categoria:
    cat = i.split(".")
    cat_pass = []
    for _ in cat:
        if not _.isnumeric():
            cat_pass.append(_)

    categoria_str.append(cat_pass)

categoria = [i[0] for i in categoria_str]
categoria = pd.Series(categoria)
#categoria

last_time = raw.iloc[3, 1:]
num_month = int(len(last_time)/4)
last_time = pd.DataFrame(last_time)
last_time.columns = ["mes"]
last_time.reset_index()
last_time = last_time.query("mes == 'janeiro 2020'")

last_time = last_time.iloc[1, :]
last_time = last_time.name
last_time = int(last_time.split(" ")[1])
#last_time, num_month

time = raw.iloc[3, 1:last_time]
time = time.reset_index(drop=True)
#time

ipca_mes = raw.iloc[4:(last_time - 1), 1:last_time]
ipca_mes = ipca_mes.reset_index(drop=True)
ipca_mes.columns = time
#ipca_mes

ipca_ano = raw.iloc[4:(last_time - 1), last_time:(last_time + num_month)]
ipca_ano = ipca_ano.reset_index(drop=True)
ipca_ano.columns = time
#ipca_ano

ipca_12m = raw.iloc[4:(last_time - 1), (last_time + num_month):(last_time + 2 * num_month)]
ipca_12m = ipca_12m.reset_index(drop=True)
ipca_12m.columns = time
#ipca_12m

peso = raw.iloc[4:(last_time - 1), (last_time + 2 * num_month):(last_time + 3 * num_month)]
peso = peso.reset_index(drop=True)
peso.columns = time

table_peso = pd.concat([categoria, peso], axis=1)
table_peso = table_peso.iloc[1:,:]
table_peso = table_peso.rename(columns={0: "Categoria"})
#table_peso

ipca_mes_last = ipca_mes.iloc[:, -1]
ipca_ano_last = ipca_ano.iloc[:, -1]
ipca_12m_last = ipca_12m.iloc[:, -1]
table_main = pd.concat([categoria, ipca_mes_last, ipca_ano_last, ipca_12m_last], axis=1)
table_main.columns = ["", ipca_mes_last.name, "Acumulado no ano", "Acumulado 12 meses"]
#table_main

URL_HIST = "https://ftp.ibge.gov.br/Precos_Indices_de_Precos_ao_Consumidor/IPCA/Serie_Historica/ipca_SerieHist.zip"

hist_raw = requests.get(URL_HIST, stream=True)
hist_raw = zipfile.ZipFile(io.BytesIO(hist_raw.content))

hist_raw_filename = str(hist_raw.filelist[0])
hist_raw_filename = hist_raw_filename.split("'")[1]

hist_raw = hist_raw.read(hist_raw_filename)
hist_raw = pd.ExcelFile(hist_raw)
hist_raw = pd.read_excel(hist_raw)
#hist_raw

hist = hist_raw.iloc[7:len(hist_raw), [0,1,3]]
hist = hist.dropna(how="all")

hist.iloc[:,0] = hist.iloc[:,0].fillna(method="ffill")
hist = hist.dropna()
hist.columns = ["ano", "mes", "valor"]
hist = hist[hist.valor != "(%)"]
hist = hist.reset_index(drop=True)
#hist

data_str = [str(hist.iloc[i,1]) + "-" + str(hist.iloc[i, 0]) for i in range(len(hist.mes))]
data_str = pd.Series(data_str)
#data_str

table_hist = pd.concat([data_str, hist.iloc[:,2]], axis=1)
table_hist.columns = ["Data", "IPCA no Mês"]
#table_hist

with pd.ExcelWriter("RENAN STELLET IPCA.xlsx") as writer:
    table_main.to_excel(writer, index=False, sheet_name="Tabela")
    table_hist.to_excel(writer, index=False, sheet_name="IPCA mensal histórico")
    table_peso.to_excel(writer, index=False, sheet_name="Pesos")

workbook = openpyxl.load_workbook("RENAN STELLET IPCA.xlsx")

xcl_main = workbook["Tabela"]
xcl_main.column_dimensions["A"].width = 24
xcl_main.column_dimensions["B"].width = 10
xcl_main.column_dimensions["C"].width = 15
xcl_main.column_dimensions["D"].width = 15
xcl_main.row_dimensions[1].height = 30

xcl_main['B1'].alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center",
 wrapText=True)
xcl_main['C1'].alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center",
 wrapText=True)
xcl_main['D1'].alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center",
 wrapText=True)

xcl_peso = workbook["Pesos"]
xcl_peso.column_dimensions["A"].width = 25
xcl_peso_columntofit = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O"]
for i in xcl_peso_columntofit:
    xcl_peso.column_dimensions[i].width = 15

xcl_hist = workbook["IPCA mensal histórico"]
xcl_hist.column_dimensions["A"].width = 12
xcl_hist.column_dimensions["B"].width = 12

workbook.save("RENAN STELLET IPCA.xlsx")
