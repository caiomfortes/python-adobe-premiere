#%% importação
import pymiere
import pandas as pd
import os

#%% organizar nomes das pastas e subpastas 
# localização da planilha que direciona o formato das pastas
path_data = 'pastas.xlsx'

# Criação de um dicionário para facilitar manipulação e acesso no código. Cada elemento do dicionário representa uma pasta.
# As subpastas serão adicionadas em um array, já que cada pasta pode ter múltiplas subpastas.
df_dict = {
    "1 - LIBERAÇÃO": [],
    "2 - ATIVAÇÕES E CORRETIVOS": [],
    "3 - ALONGAMENTOS": [],
    "4 - POTÊNCIA": [],
    "5 - CORE": [],
    "6 - FORÇA - SUPERIORES": [],
    "7 - FORÇA - INFERIORES": [],
}

# Abre a planilha que possui a estrutura das pastas no formato exemplificado
df = pd.read_excel(path_data, sheet_name="Sheet1")

# Percorre a planilha e adiciona as subpastas em cada pasta correspondente no dicionário.
for i in range(0,df.shape[0],1):
    if(i == 0):
        df_dict['1 - LIBERAÇÃO'].append(df['SUB-PASTA'][i])
    
    elif(df['PASTA'][i] ==  "1 - LIBERAÇÃO" and df['SUB-PASTA'][i] != df['SUB-PASTA'][i-1]):
        df_dict['1 - LIBERAÇÃO'].append(df['SUB-PASTA'][i])
        
    elif(df['PASTA'][i] ==  "2 - ATIVAÇÕES E CORRETIVOS" and df['SUB-PASTA'][i] != df['SUB-PASTA'][i-1]):
        df_dict["2 - ATIVAÇÕES E CORRETIVOS"].append(df['SUB-PASTA'][i])
    
    elif(df['PASTA'][i] ==  "3 - ALONGAMENTOS" and df['SUB-PASTA'][i] != df['SUB-PASTA'][i-1]):
        df_dict["3 - ALONGAMENTOS"].append(df['SUB-PASTA'][i])
    
    elif(df['PASTA'][i] ==  "4 - POTÊNCIA" and df['SUB-PASTA'][i] != df['SUB-PASTA'][i-1]):
        df_dict["4 - POTÊNCIA"].append(df['SUB-PASTA'][i])
   
    elif(df['PASTA'][i] ==  "5 - CORE" and df['SUB-PASTA'][i] != df['SUB-PASTA'][i-1]):
        df_dict["5 - CORE"].append(df['SUB-PASTA'][i])
    
    elif(df['PASTA'][i] ==  "6 - FORÇA - SUPERIORES" and df['SUB-PASTA'][i] != df['SUB-PASTA'][i-1]):
        df_dict["6 - FORÇA - SUPERIORES"].append(df['SUB-PASTA'][i])
   
    elif(df['PASTA'][i] ==  "7 - FORÇA - INFERIORES" and df['SUB-PASTA'][i] != df['SUB-PASTA'][i-1]):
        df_dict["7 - FORÇA - INFERIORES"].append(df['SUB-PASTA'][i])


#%% criação das pastas no premiere

# como o projeto do premiere escolhido aberto, esse comando cria uma pasta e suas subpastas.
# A primeira pasta a ser criada é sempre o nome do próprio projeto, para melhorar a organização, mas isso pode ser alterado.
# É interessante que apenas um projeto esteja aberto para evitar conflitos.

# nome de todos os projetos
names = ["1 - LIBERAÇÃO","2 - ATIVAÇÕES E CORRETIVOS","3 - ALONGAMENTOS","4 - POTÊNCIA","5 - CORE","6 - FORÇA - SUPERIORES","7 - FORÇA - INFERIORES"]
# número correspondente ao projeto aberto e escolhido para manipulação no array names
proj = 5    # "6 - FORÇA - SUPERIORES"

# executa a criação da pasta inicial com o nome do projeto
pymiere.objects.app.project.rootItem.createBin(names[proj])

# percorre o elemento do dicionário com o nome do projeto criando as subpastas presentes
for j in range(0,len(df_dict[names[proj]]),1):
    
    # .children[0] sinaliza a primeira pasta/elemento criado no premiere. Para o caso de ser o segundo (se criasse algo antes do código),
    # seria usado .children[1]. Essa dinâmica serve para subpastas tbm, a primeira subpasta criada começará em 0.
    # aqui é criada uma subpasta dentro da primeira pasta criada (que possoi o mesmo nome do projeto)
    pymiere.objects.app.project.rootItem.children[0].createBin(df_dict[names[proj]][j])
    
    # aqui ocorre a criação de duas subpastas dentro de cada subpasta anterior
    pymiere.objects.app.project.rootItem.children[0].children[j].createBin("CAM 1")
    pymiere.objects.app.project.rootItem.children[0].children[j].createBin("CAM 2")
    
    # pela ordem, caso precisasse acessar acessar as pastas CAM 1 e CAM 2, poderia ser feito o seguinte código:
    # pymiere.objects.app.project.rootItem.children[0].children[j].children[0] - acessa CAM 1, a primeira a ser criada
    # pymiere.objects.app.project.rootItem.children[0].children[j].children[1] - acessa CAM 2, a segunda a ser criada

#%%
# aqui ocorre a criação das sequências dentro do projeto nomeado elas conforme um padrão escolhido

# localização do arquivo excel que contém o nome de cada exercício.
# Este código pode facilmente ser alterado e simplificado, pois para o caso em específico, devido
# a grande quantidade de vídeos de diferentes tipo, é necessário uma estrutura de nomenclatura mais complexa.
# Tudo isso foi definido de forma rápida em arquivo excel anteriormente
# Estrutura do nome desejada: Número do exercício - Nome do exercício - Grupo do exercício (subpasta) - Categoria do exercício (projeto)

# nome de todos os projetos
names = ["1 - LIBERAÇÃO","2 - ATIVAÇÕES E CORRETIVOS","3 - ALONGAMENTOS","4 - POTÊNCIA","5 - CORE","6 - FORÇA - SUPERIORES","7 - FORÇA - INFERIORES"]

# numero do projeto escolhido (categoria)
proj = 5
path_data2 = r"BANCO DE EXERCÍCIOS.xlsx" # não pode ser disponibilizada publicamente - cada aba da planilha representa um projeto, ou seja, uma categoria
df2 = pd.read_excel(path_data2, sheet_name=str(proj+1))

# dentro de cada aba de categoria na planilha, possuia seus exercícios, numeros, grupos e a categoria separados em células
# aqui é definido o grupo que irá começar, ou seja, o primeiro que aparece
nome_grupo = df2['GRUPO'][0]

# voltamos a percorrer o dicionario criado anteriormente no projeto escolhido
count = 0
for j in range(0,len(df_dict[names[proj]]),1):
    
    # percorremos cada exercício e paramos quando criarmos todos os exercício pertencentes a uma subpasta, e então pula para a próxima
    # quando não houverem mais exercícios na lista, será encerrado
    x=0
    while (x<1):
        
        # nome da sequência garantindo conversão para str
        sequence_name = str(df2['NÚMERO'][count]) + ' - '  + df2['NOME'][count] + ' - ' + df2['GRUPO'][count] + ' - ' + df2['CATEGORIA'][count]
        
        # criação das sequência nas subpasta (Estrutura das pastas: Projeto/Categoria -> Grupo do exercício -> CAM 1 + CAM 2 + Sequências de cada exercício)
        # [pymiere.objects.app.project.rootItem.children[1]] este comando representa uma sequêcia base que foi criada
        # com as configurações desejadas para todas as sequências. Ela foi criada posterior a criação das postas, por isso é numerada como children[1]
        pymiere.objects.app.project.createNewSequenceFromClips(sequence_name,[pymiere.objects.app.project.rootItem.children[1]],pymiere.objects.app.project.rootItem.children[0].children[j])
        
        count = count + 1
        if(df2["GRUPO"][count] == ""):
            break
        # Já que uma aba da planilha possui vários grupos de exercícios, se o grupo muda na leitura, é necessário mudar a pasta
        # O x encerra o while, é mudado o nome do grupo na variável e a leitura do dicionário avança para o próximo grupo
        elif(nome_grupo != df2["GRUPO"][count]):
            nome_grupo = df2["GRUPO"][count]
            x = x+1
            
            
#%% importação dos vídeos dentro de cada pasta CAM 1 e CAM 2
import os
import glob

# nome de todos os projetos
names = ["1 - LIBERAÇÃO","2 - ATIVAÇÕES E CORRETIVOS","3 - ALONGAMENTOS","4 - POTÊNCIA","5 - CORE","6 - FORÇA - SUPERIORES","7 - FORÇA - INFERIORES"]

# numero do projeto escolhido (categoria)
proj = 5

# Define o caminho da pasta com os vídeos (neste caso, tinha pastas com os nomes dos projetos/grupos e seus vídeos dentro)
path_videos = r"GRAVAÇÕES/"
path_videos = path_videos + names[proj] + "/"

#importa os vídeos
# percorre o dicionário anterior no projeto escolhido
for i in range(0,len(df_dict[names[proj]]),1):
    # define o caminho das pastas cam1 e cam2
    path_videos_cam1 = path_videos + df_dict[names[proj]][i] + "/CAM 1"
    path_videos_cam2 = path_videos + df_dict[names[proj]][i] + "/CAM 2"
    
    # seleciona todos os vídeo .mp4 da pasta cam1
    videos_cam1 = glob.glob(os.path.join(path_videos_cam1, "*.mp4"))
    for j in range(len(videos_cam1)-1,0,-1):
        arquivo = videos_cam1[j]
        
        nome_arquivo, extensao = arquivo.split(".")
        ultimos_digitos = nome_arquivo[-3:]
        # os vídeos terminados em S03 antes do .mp4 eram uma identificação que a camera usava para os proxies (videos iguais em menor qualidade)
        # importante para quem trabalha com proxies. Eles serão associados logo abaixo com seus vídeos correspondentes
        if(ultimos_digitos == 'S03'):
            del videos_cam1[j]


    # repeto o processi para a pasta cam2
    videos_cam2 = glob.glob(os.path.join(path_videos_cam2, "*.mp4"))
    for j in range(len(videos_cam2)-1,0,-1):
        arquivo = videos_cam2[j]
        nome_arquivo, extensao = arquivo.split(".")
        ultimos_digitos = nome_arquivo[-3:]
        if(ultimos_digitos == 'S03'):
            del videos_cam2[j]
    
    # importa todos os vídeos coletados para suas pastas no premiere
    # children[i] - projeto/categoria | children[i] - grupo | children[0] - CAM 1
    pymiere.objects.app.project.importFiles(videos_cam1,True, pymiere.objects.app.project.rootItem.children[0].children[i].children[0],False)
    pymiere.objects.app.project.importFiles(videos_cam2,True, pymiere.objects.app.project.rootItem.children[0].children[i].children[1],False)

    # coleta todos os proxies (videos que não foram coletados anteriormente)
    proxy_cam1 = glob.glob(os.path.join(path_videos_cam1, "*.mp4"))
    proxy_cam1 = [x for x in proxy_cam1 if x not in videos_cam1]
    
    #criar proxies
    # esse processo considera que os vídeos foram coletados em ordem alfabética, assim como a sua importação foi em ordem alfabética
    # ou seja, a ordem da lista de vídeos e da pasta seria como [0001.mp4, 0002.mp4, 0003.mp4], e seriam importados nessa msm ordem
    # e a ordem da lista de proxies seria [0001S03.mp4, 0002S03.mp4, 0003S03.mp4]
    x = 0
    while(x<len(videos_cam1)):
        for j in videos_cam1:
            # verifica correspondência entre o proxy e o seu video original para confirmar que o vídeo possui aquele proxy e faz a associação,
            # apenas para e evitar erros de anexo.
            # Este código se mantém incompleto, pois considera que todos os vídeos possuem proxy, é necessária uma melhoria.
            # Mas se um arquivo no meio de outros não possuir proxy, provavelmente
            # teriamos uma quebra de ritmo e os próximos ficariam sem anexar os seus proxies. Também existem outro possíveis erros
            # Analise o seu caso e ajuste o código se pretender usá-lo
            nome_arquivo, extensao = videos_cam1[x].split(".")
            nome_arquivo_proxy, extensao = proxy_cam1[x].split(".")
            if(nome_arquivo_proxy == (nome_arquivo + "S03")):
                a = pymiere.objects.app.project.rootItem.children[0].children[i].children[0].children[x].getMediaPath()
                caminho, arq = a.split('C0')
                num_arq = nome_arquivo[-5:]
                pymiere.objects.app.project.rootItem.children[0].children[i].children[0].children[x].attachProxy(caminho +num_arq+'S03.MP4', 0)
                x = x+1
                break
            else:
                x = x + 1
            
    
    proxy_cam2 = glob.glob(os.path.join(path_videos_cam2, "*.mp4"))
    proxy_cam2 = [x for x in proxy_cam2 if x not in videos_cam2]
    x = 0
    while(x<len(videos_cam2)):
        for j in videos_cam2:
            nome_arquivo, extensao = videos_cam2[x].split(".")
            nome_arquivo_proxy, extensao = proxy_cam2[x].split(".")
            if(nome_arquivo_proxy == (nome_arquivo + "S03")):
                a = pymiere.objects.app.project.rootItem.children[0].children[i].children[1].children[x].getMediaPath()
                caminho, arq = a.split('C0')
                num_arq = nome_arquivo[-5:]
                pymiere.objects.app.project.rootItem.children[0].children[i].children[1].children[x].attachProxy(caminho +num_arq+'S03.MP4', 0)
                x = x+1
                break
            else:
                x = x + 1

# %%
