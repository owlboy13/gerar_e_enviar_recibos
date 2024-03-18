import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
import win32com.client as win32
import os
from tqdm.auto import tqdm


name_planilha = input('Digite o nome da planilha aqui e aperte ENTER:  ')
format_xslx = '.xlsx'
#lendo arquivo excel
caminho_planilha = name_planilha + format_xslx 

planilha_excel = pd.read_excel(caminho_planilha)
print()
#caminho do diretorio dos recibos.pdfs
save_recibo = input('Cole aqui o endereço da pasta que deseja salvar os recibos e aperte ENTER:  ')
if os.path.isdir(save_recibo):
    caminho_recibo = save_recibo
else:
    print("O caminho fornecido não é uma pasta válida.")
print()
user_franquia = input(str('Qual a Franquia?  ')).upper()
print()


for index, row in tqdm(planilha_excel.iterrows()):

    #nomeando cada recibo com nome que esta na coluna da
    name_file = row['Nome']

    #gerando recibo em pdf
    pdf_file = canvas.Canvas(f'{name_file}.pdf')

    fonte_padrao = 'Helvetica-Bold'
    tamanho_padrao = 12

    pdf_file.setFont(fonte_padrao, tamanho_padrao)

     # Acessando as informações da planilha
    pdf_file.drawString(200, 650,f"Recibo de {row['Recibo']}")
    pdf_file.drawString(50, 450, f"Nome: {row['Nome']}")
    pdf_file.drawString(50, 425, f"CNPJ: {row['CNPJ']}")
    pdf_file.drawString(50, 400, f"Período: {row['Período']}")
    valor_sem_decimal = int(row['Valor'])  # Converter o valor para um número int
    pdf_file.drawString(50, 375, f"Valor: R$ {valor_sem_decimal:,},00")

    tamanho_fonte = 10
    pdf_file.setFont(fonte_padrao, tamanho_fonte)

    #inserir logo e assinatura bronx
    if user_franquia == 'BRONX':
        pdf_file.drawImage("logo_bronx_cabecalho.png", 0, 700, width=2*inch, height=2*inch)
        pdf_file.drawImage("cabeçalho_lado_direito_bronx.png", 305, 700, width=4*inch, height=2*inch)
        pdf_file.drawImage("assinatura_entregador.png", 350, 150, width=2*inch, height=0.5*inch)
        pdf_file.drawImage("assinatura_bronx.jpg", 60, 140, width=3*inch, height=2*inch)

        #declaração bronx
        pdf_file.drawString(20, 600, f"O prestador {row['Nome']} CNPJ {row['CNPJ']}, declara ter recebido de")
        pdf_file.drawString(20, 580, f"BRONX INCORPORACAO LOGISTICA LTDA CNPJ 48.834.064/0001-14 o valor de R$ {valor_sem_decimal:,},00 referente a")
        pdf_file.drawString(20, 560, "Gorjetas no período citado abaixo:")

    #assinatura e logo jafhelog
    elif user_franquia == 'JAFHELOG':
        pdf_file.drawImage("logo_jafhelog_cabecalho.png", 0, 700, width=2*inch, height=2*inch)
        pdf_file.drawImage("cabeçalho_lado_direito_jafhelog.png", 305, 700, width=3*inch, height=1.5*inch)
        pdf_file.drawImage("assinatura_entregador.png", 350, 150, width=2*inch, height=0.5*inch)
        pdf_file.drawImage("assinatura_jafhe.png", 60, 140, width=3*inch, height=2*inch)

        # decaração jafhelog
        pdf_file.drawString(22, 600, f"O prestador {row['Nome']} de CNPJ {row['CNPJ']}, declara ter recebido de")
        pdf_file.drawString(22, 580, f"JAFHE LOG INCORPORACAO LOGISTICA LTDA CNPJ 50.685.447/0001-10 o valor de R$ {valor_sem_decimal:,},00 referente a")
        pdf_file.drawString(22, 560, "Gorjetas no período citado abaixo:")
    
    else:
        print('Franquia não existe')
    

    pdf_file.save()

if user_franquia == 'BRONX':
    rodape = r'https://i.imgur.com/EIoVMgm.png'
elif user_franquia == 'JAFHELOG':
    rodape = r'https://i.imgur.com/kkEfrhY.png'


#ENVIO DE EMAILS
print()
print('Antes de confirmar o envio verifique se os recibos estão corretos')
print()

periodo_email = input(str('Qual o periodo referente as gorjetas?  '))
print()
autoriz_send = input('Posso enviar os e-mails?  ').lower()
print()


if autoriz_send == 'sim' or autoriz_send == 's':
    for indice, row in tqdm(planilha_excel.iterrows()):

        #lendo email na planilha
        email_receiver = row['email']
        #acessando outlook para enviar email
        outlook = win32.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)
        message.To = email_receiver
            
        # adicionando indice para seguir sequencia dos emails na planilha
        name_excel = row['Nome']
        pdf_files = (f'{name_excel}.pdf')

        pdf_path = os.path.join(caminho_recibo, pdf_files)
        #conteudo do email
        if os.path.exists(pdf_path):
            message.Subject = f"Recibo de Pagamento de Gorjetas - {user_franquia} "
            message.BodyFormat = 2
            message.HTMLBody = f'''<p>Olá Parceiro, tudo bem?</p>
            <p>Segue o Recibo Referente às Gorjetas do Período de {periodo_email}</p>    
            <p>Equipe {user_franquia}</p>
            <p>Agradecemos sua atenção contínua e confiança em nossa empresa. 🤝 </p>
            <p>Recibo em Anexo 📎</p>
            <p><i>Este e-mail é automático e não é necessário respondê-lo</i></p>
            <img src="{rodape}" alt="Imagem" width="210" height="100">'''
            #anexando recibo pdf
            attachment = message.Attachments.Add(pdf_path)
            message.Send()
        else:
            print(f'Arquivo PDF Não encontrado para {email_receiver}: {pdf_file}')
else:
    print('Algum problema nos recibos?')
print()
print('Abra o Outlook e aguarde os emails sairem da caixa de saída')
print()
input("Quando o Outlook finalizar, aperte ENTER...")
