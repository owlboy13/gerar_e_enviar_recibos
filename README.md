# gerar_e_enviar_recibos.py

# RESUMO
A ideia desse código é gerar arquivos em formato pdfs baseado em informações de uma planilha, como nome, cnpj, data, valor, logomarca do contratante e assinatura.
Depois de gerar os arquivos, enviar esses mesmos arquivos por email(outlook) de forma sequencial baseado tambem em uma lista de emails na mesma planilha.

O código é divido em duas partes:
1- A primeira parte consiste em gerar arquivos através das informações de uma planilha com os dados pessoais dos individuos que irão receber o e-mail, o código busca essas informações e gera os arquivos em massa dentro da pasta onde fica o código.
2- A segunda parte é o envio desses arquivos, o código busca os endereços de emails na planilha seguindo a mesma sequencia das informações pessoais dos destinatarios e através do modulo PyWin32Com conectar ao outlook e fazer o envio por essa ferramente.

Dessa forma conseguimos otimizar o tempo que levaria para enviar recibos em massa para um numero alto de destinatarios.

O código pode ser usado para enviar qualquer outro tipo de arquivo e informação, por exemplo:

- envio de imagens de marketing e captação de clientes;
- envio de contratos em massa;
- envio de boletos;
- lembretes e avisos
