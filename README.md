Fluxograma de Execução da Aplicação
1. Início: O usuário abre a aplicação.
2. Interface Gráfica: A interface gráfica é exibida, com um botão "Buscar Previsão".
3. Ação do Usuário: O usuário clica no botão "Buscar Previsão".
4. Conexão à Internet: A aplicação, usando a biblioteca requests, faz uma
requisição para a API de previsão do tempo.
5. Captura de Dados: A aplicação recebe os dados em formato JSON da API e extrai
a temperatura e o status da umidade do ar.
6. Verificação do Arquivo: A aplicação verifica se o arquivo da planilha
(dados_previsao.xlsx) existe no mesmo diretório.
7. Manipulação da Planilha:
○ Se o arquivo não existir: Um novo arquivo de planilha é criado, com os
cabeçalhos: "Data/Hora", "Temperatura" e "Status da Umidade do Ar".
○ Se o arquivo já existir: A planilha é aberta para adicionar novos dados.
8. Inserção dos Dados: A data e hora atuais, a temperatura e a umidade capturadas
são inseridas em uma nova linha da planilha.
9. Salvamento: O arquivo da planilha é salvo.
10. Fim: Uma mensagem de confirmação é exibida na interface para o usuário.
