# Envio Automático de E-mails com Outlook e Python  

## Descrição  
Este script utiliza a biblioteca `win32com.client` para enviar e-mails automaticamente via Microsoft Outlook. Ele carrega uma planilha Excel contendo os e-mails dos destinatários, seus respectivos nomes e os relatórios a serem enviados como anexo.  

## Como Usar  

1. **Pré-requisitos**  
   - Ter o Microsoft Outlook instalado e configurado no computador.  
   - Instalar as bibliotecas necessárias executando:  

     ```bash
     pip install pywin32 pandas
     ```  

2. **Estrutura do Arquivo Excel**  
   O script espera um arquivo Excel (`Enviar E-mails.xlsx`) com as seguintes colunas:  

   - **E-mail**: Endereço de e-mail do destinatário.  
   - **Gerente**: Nome do gerente responsável.  
   - **Relatório**: Nome do arquivo de relatório a ser enviado.  

3. **Execução do Script**  
   - Altere o caminho do arquivo Excel no código para corresponder à sua estrutura local.  
   - Altere o caminho onde os relatórios estão armazenados.  
   - Execute o script no terminal ou em um ambiente Python:  

     ```bash
     python nome_do_arquivo.py
     ```  

# Contribuindo para o projeto
1. Para contribuir com <nome_do_projeto>, siga estas etapas:
2. Bifurque este repositório.
3. Crie um branch: git checkout -b <nome_branch>.
4. Faça suas alterações e confirme-as: git commit -m '<mensagem_commit>'
5. Envie para o branch original: git push origin <nome_do_projeto>/<local>
6. Crie a solicitação de pull.

Como alternativa, consulte a documentação do GitHub em [como criar uma solicitação pull](https://help.github.com/en/github/collaborating-with-issues-and-pull-requests/creating-a-pull-request).

