import os
import win32com.client as win32
import pandas as pd
from datetime import datetime
from collections import defaultdict

# Caminhos
planilha_caminho = r"C:\Users\trank\Desktop\Pasta1.xlsx"
historico_path = r"historico_notificacoes.xlsx"

# Verifica se a planilha existe
print(f"Verificando se a planilha mensal existe: {planilha_caminho}")
if not os.path.exists(planilha_caminho):
    print("❌ ERRO: Planilha não encontrada.")
    exit()

# Lê os dados da planilha principal
df = pd.read_excel(planilha_caminho)

# Mostra as colunas disponíveis para evitar futuros erros
print("📊 Colunas encontradas na planilha:", df.columns.tolist())

# Lê o histórico ou cria um novo DataFrame se não existir
if os.path.exists(historico_path):
    historico_df = pd.read_excel(historico_path)
else:
    historico_df = pd.DataFrame(columns=["Nome", "Email", "Data"])

# Agrupa as datas esquecidas por funcionário
funcionarios = defaultdict(list)

for _, row in df.iterrows():
    nome = row['Nome']
    email = row['Email']
    data = row['Data']
    funcionarios[(nome.strip().lower(), email.strip().lower())].append(str(data))

# Conecta ao Outlook
outlook = win32.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Lista as contas disponíveis
print("📬 Contas disponíveis no Outlook:")
for i in range(namespace.Accounts.Count):
    conta = namespace.Accounts.Item(i + 1)
    print(f"- {conta.DisplayName} | {conta.SmtpAddress}")

# Define o e-mail de envio
email_envio = "fisioap@outlook.com.br"
conta_encontrada = None

for i in range(namespace.Accounts.Count):
    conta = namespace.Accounts.Item(i + 1)
    if conta.SmtpAddress.lower() == email_envio.lower():
        conta_encontrada = conta
        break

if not conta_encontrada:
    print(f"❌ ERRO: Conta de e-mail '{email_envio}' não encontrada no Outlook.")
    exit()

print("📧 Conexão com o Outlook estabelecida com sucesso!")

# Envia os e-mails
for (nome, email), datas in funcionarios.items():
    # Verifica histórico
    historico = historico_df[(historico_df['Email'].str.lower() == email)]
    notificacoes_anteriores = historico.shape[0]
    notificacao_atual = notificacoes_anteriores + 1

    print(f"DEBUG: Nome: {nome} | Email: {email} | Notificações Anteriores: [{notificacoes_anteriores}] | Notificação Atual: {notificacao_atual}")

    # Gera assunto com número ordinal
    ordinais = {1: "1ª", 2: "2ª", 3: "3ª", 4: "4ª", 5: "5ª", 6: "6ª", 7: "7ª", 8: "8ª", 9: "9ª", 10: "10ª"}
    prefixo = ordinais.get(notificacao_atual, f"{notificacao_atual}ª")
    assunto = f"{prefixo} Notificação de Esquecimento"
    print(f"DEBUG: Assunto do e-mail antes do envio -> {assunto}")

    corpo = f"""
Olá {nome.title()},

Identificamos que você esqueceu de registrar o ponto nas seguintes datas:

{chr(10).join(f"- {d}" for d in datas)}

Essa é sua {prefixo.lower()} notificação sobre esse tipo de ocorrência. Por favor, redobre a atenção para evitar impactos no controle de frequência.

Atenciosamente,
Recursos Humanos
"""

    # Cria o e-mail
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = assunto
    mail.Body = corpo
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, conta_encontrada))  # Define a conta de envio

    try:
        mail.Send()
        print(f"✅ Conta de envio definida como: {conta_encontrada.SmtpAddress}")
        print(f"✅ E-mail enviado para {email}")

        # Atualiza histórico
        for d in datas:
            historico_df = pd.concat([historico_df, pd.DataFrame([{'Nome': nome, 'Email': email, 'Data': d}])], ignore_index=True)

    except Exception as e:
        print(f"❌ ERRO ao enviar e-mail para {email}: {e}")

# Salva o histórico
historico_df.to_excel(historico_path, index=False)
print("✅ Notificações enviadas e histórico atualizado!")
