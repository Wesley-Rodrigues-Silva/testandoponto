import pandas as pd
import win32com.client
import os

# 🚀 Mensagem de depuração para garantir que o código atualizado está rodando
print("\n🚀 ESTE É O CÓDIGO ATUALIZADO RODANDO! 🚀\n")

# 🛠 Caminhos corretos para os arquivos
planilha_mensal = r"C:\Users\trank\Desktop\Pasta1.xlsx"
historico_notificacoes = r"C:\Users\trank\Desktop\historico.xlsx"

# 📌 Depuração: Exibir os caminhos usados
print("Executando o script...")
print(f"Verificando se a planilha mensal existe: {planilha_mensal}")

# 🔍 Verificar se o arquivo da planilha existe
if not os.path.exists(planilha_mensal):
    print(f"❌ ERRO: Arquivo {planilha_mensal} não encontrado!")
    exit()

# 🔍 Verificar se o histórico de notificações existe
if os.path.exists(historico_notificacoes):
    historico_df = pd.read_excel(historico_notificacoes, engine="openpyxl")
else:
    historico_df = pd.DataFrame(columns=["Nome", "Email", "Notificacoes", "Ultimas Datas"])

# 📥 Carregar os dados da planilha mensal
mensal_df = pd.read_excel(planilha_mensal, engine="openpyxl")

# 📌 Converter datas para o formato brasileiro (se a coluna existir)
if "Data" in mensal_df.columns:
    mensal_df["Data"] = pd.to_datetime(mensal_df["Data"], errors='coerce').dt.strftime('%d/%m/%Y')
else:
    print("❌ ERRO: A coluna 'Data' não foi encontrada na planilha.")
    exit()

# 📌 Criar um dicionário para armazenar notificações
notificacoes = {}
for _, row in mensal_df.iterrows():
    nome, email, data = row["Nome"], row["Email"], row["Data"]
    if email in notificacoes:
        notificacoes[email]["datas"].append(data)
    else:
        notificacoes[email] = {"nome": nome, "datas": [data]}

# 📧 Inicializar o Outlook para envio de e-mails
try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    print("📧 Conexão com o Outlook estabelecida com sucesso!")
except Exception as e:
    print(f"❌ ERRO ao conectar ao Outlook: {e}")
    exit()

# 📌 Selecionar a conta específica para envio
conta = None
email_envio = "fisioap@outlook.com.br"  # Altere para o e-mail correto

for acc in outlook.Session.Accounts:
    if acc.SmtpAddress.lower() == email_envio.lower():
        conta = acc
        break

if not conta:
    print(f"❌ ERRO: Conta de e-mail '{email_envio}' não encontrada no Outlook.")
    exit()

# 📤 Iterar sobre os funcionários a serem notificados
for email, info in notificacoes.items():
    nome = info["nome"]
    datas = info["datas"]
    
    # 📌 Verificar quantas notificações o funcionário já recebeu
    notificacoes_anteriores = historico_df.loc[historico_df["Email"] == email, "Notificacoes"].values
    num_notificacoes = int(notificacoes_anteriores[0]) + 1 if len(notificacoes_anteriores) > 0 else 1

    # 📌 Mensagem de debug para verificar os valores antes do envio
    print(f"DEBUG: Nome: {nome} | Email: {email} | Notificações Anteriores: {notificacoes_anteriores} | Notificação Atual: {num_notificacoes}")

    # 📌 Criar o corpo do e-mail
    texto_datas = f"você esqueceu de registrar seu ponto na seguinte data: {datas[0]}" if len(datas) == 1 \
        else f"você esqueceu de registrar seu ponto nas seguintes datas: {', '.join(datas)}"
    mensagem = (f"Prezado(a) {nome},\n\n"
                f"Identificamos que {texto_datas}.\n\n"
                "Atenciosamente,\nRecursos Humanos")

    # 📌 Definir o assunto do e-mail corretamente
    assunto_email = f"{num_notificacoes}ª Notificação de Esquecimento"
    print(f"DEBUG: Assunto do e-mail antes do envio -> {assunto_email}")  # Verifica se o assunto está correto

    # 📧 Criar e enviar o e-mail via Outlook
    try:
        mail = outlook.CreateItem(0)
        mail.SendUsingAccount = conta  # Definir a conta correta para envio
        mail.To = email
        mail.CC = "chefia@exemplo.com"  # E-mail da chefia (adicionar no CC)
        mail.Subject = assunto_email  # Aplicar o assunto corretamente
        mail.Body = mensagem  

        # Exibir o e-mail antes de enviar para depuração
        mail.Display()  # Abrir o e-mail para depuração
        
        # Enviar o e-mail (remova ou comente se quiser apenas visualizar o e-mail)
        # mail.Send()  
        print(f"✅ E-mail preparado para {email} com o assunto: {mail.Subject}")
    except Exception as e:
        print(f"❌ ERRO ao enviar e-mail para {email}: {e}")
    
    # 📌 Atualizar o histórico de notificações
    historico_df = historico_df[historico_df["Email"] != email]
    historico_df = pd.concat([historico_df, pd.DataFrame([{
        "Nome": nome,
        "Email": email,
        "Notificacoes": num_notificacoes,
        "Ultimas Datas": ', '.join(datas)
    }])], ignore_index=True)

# 💾 Salvar o histórico atualizado
historico_df.to_excel(historico_notificacoes, index=False, engine="openpyxl")
print("✅ Notificações enviadas e histórico atualizado!")
