import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
import win32com.client as win32
import os
from tqdm.auto import tqdm

#função para reiniciar programa
def main():
    def printbar():
        print('///////////////////////////////////////////')
    name_planilha = input('Digite o nome da planilha aqui e aperte ENTER:  ')
    format_xslx = '.xlsx'
    #lendo arquivo excel
    caminho_planilha = name_planilha + format_xslx 
    planilha_excel = pd.read_excel(caminho_planilha)
    printbar()
    #caminho do diretorio dos recibos.pdfs
    save_recibo = input(r'Cole aqui o endereço da pasta que deseja salvar os recibos e aperte ENTER:  ')
    try:
        os.path.isdir(save_recibo)
        caminho_recibo = save_recibo
    except:
        print("O caminho fornecido não é uma pasta válida.")
        erro_caminho = input('Quer tentar novamente?  ')
        if erro_caminho == 'sim' or erro_caminho == 's':
            main()
        else:
            printbar()
            input('Aperte enter para sair... ')

    printbar()
    user_franquia = input(str('Qual a Franquia?  ')).upper()
    printbar()

    # Referencia do pagamento 
    referencia = input('Digite a referencia -> Recibo de:  ')

    for indice, row in tqdm(planilha_excel.iterrows()):

        #nomeando cada recibo com nome que esta na coluna da
        name_file = row['Nome']

        #gerando recibo em pdf
        pdf_file = canvas.Canvas(os.path.join(caminho_recibo, f'{name_file}.pdf'))

        fonte_padrao = 'Helvetica-Bold'
        tamanho_padrao = 12

        pdf_file.setFont(fonte_padrao, tamanho_padrao)

        # Acessando as informações da planilha
        pdf_file.drawString(200, 650,f"Recibo de {row['Recibo']}")
        pdf_file.drawString(50, 450, f"Nome: {row['Nome']}")
        pdf_file.drawString(50, 425, f"CNPJ: {row['CNPJ']}")
        pdf_file.drawString(50, 400, f"Período: {row['Período']}")
        valor_float = float(row['Valor']) # Converter o valor para um número int
        valor_str = f"{valor_float:,.2f}".replace('.',',') #valor com casa decimal ,00
        pdf_file.drawString(50, 375, f"Valor: R$ {valor_str}")

        tamanho_fonte = 10
        pdf_file.setFont(fonte_padrao, tamanho_fonte)

        #inserir logo e assinatura bronx
        if user_franquia == 'BRONX':
            pdf_file.drawImage("logo_bronx_cabecalho.png", 0, 700, width=2*inch, height=2*inch)
            pdf_file.drawImage("cabeçalho_lado_direito_bronx.png", 305, 700, width=3.8*inch, height=1.8*inch)
            pdf_file.drawImage("assinatura_entregador.png", 350, 150, width=2*inch, height=0.5*inch)
            pdf_file.drawImage("assinatura_bronx.jpg", 60, 140, width=3*inch, height=2*inch)

            #declaração bronx
            pdf_file.drawString(20, 600, f"O prestador {row['Nome']} CNPJ {row['CNPJ']}, declara ter recebido de")
            pdf_file.drawString(20, 580, f"BRONX INCORPORACAO LOGISTICA LTDA CNPJ 48.834.064/0001-14 o valor de R$ {valor_str} referente a")
            pdf_file.drawString(20, 560, f"{referencia} no período citado abaixo:")

        #assinatura e logo jafhelog
        elif user_franquia == 'JAFHELOG':
            pdf_file.drawImage("logo_jafhelog_cabecalho.png", 0, 700, width=2*inch, height=2*inch)
            pdf_file.drawImage("cabeçalho_lado_direito_jafhelog.png", 305, 700, width=3.8*inch, height=1.8*inch)
            pdf_file.drawImage("assinatura_entregador.png", 350, 150, width=2*inch, height=0.5*inch)
            pdf_file.drawImage("assinatura_jafhe.png", 60, 140, width=3*inch, height=2*inch)

            # decaração jafhelog
            pdf_file.drawString(22, 600, f"O prestador {row['Nome']} de CNPJ {row['CNPJ']}, declara ter recebido de")
            pdf_file.drawString(22, 580, f"JAFHE LOG INCORPORACAO LOGISTICA LTDA CNPJ 50.685.447/0001-10 o valor de R$ {valor_str} referente a")
            pdf_file.drawString(22, 560, f"{referencia} no período citado abaixo:")
        
        else:
            print(' - Franquia não existe!')
            print()
            erro_franquia = input('Quer tentar novamente?  ')
            if erro_franquia == 'sim' or erro_franquia == 's':
                main()
            else:
                printbar()
                input('Aperte enter para sair... ')
            

        pdf_file.save()

    if user_franquia == 'BRONX':
        rodape = r'https://i.imgur.com/EIoVMgm.png'
    elif user_franquia == 'JAFHELOG':
        rodape = r'https://i.imgur.com/kkEfrhY.png'


    #ENVIO DE EMAILS
    print()
    print('ATENÇÃO: ANTES DE ENVIAR OS EMAILS VERIFIQUE SE OS RECIBOS ESTÃO CORRETOS!')
    print()

    periodo_email = input(str('Qual o periodo referente?  '))
    printbar()
    autoriz_send = input('Posso enviar os e-mails?  ').lower()
    printbar()


    if autoriz_send == 'sim' or autoriz_send == 's':
        for index, row in tqdm(planilha_excel.iterrows()):

            #lendo email na planilha
            email_receiver = row['email']
            #acessando outlook para enviar email
            outlook = win32.Dispatch("Outlook.Application")
            message = outlook.CreateItem(0)
            message.To = email_receiver
                
            # adicionando row para seguir sequencia dos dados na planilha e renomear os arquivos
            name_excel = row['Nome']
            pdf_files = (f'{name_excel}.pdf')

            pdf_path = os.path.join(caminho_recibo, pdf_files)
            #conteudo do email   
            if os.path.exists(pdf_path):
                message.Subject = f"Recibo de {row['Recibo']} - {user_franquia} "
                message.BodyFormat = 2
                message.HTMLBody = f'''<p>Olá Parceiro, tudo bem?</p>
                <p>Segue o Recibo Referente à {referencia} do Período de {periodo_email}</p>    
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
    
        print()
        print('Abra o Outlook e aguarde os emails sairem da caixa de saída')
        print() 

    else:
        print('Algum problema nos recibos?')
        print()

    restart = input("Quer gerar e enviar os recibos novamente?  ").lower()
    if restart == 'sim' or restart == 's':
        print('Pode recomeçar:')
        print()
        main()
    else:
        printbar()
        print('Finalizado!')
if __name__ == "__main__":
    main()

input('Aperte Enter para sair... ')


