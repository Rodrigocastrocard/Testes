Documentação da Ferramenta de Cotação Intelipost
Visão geral
Esta ferramenta é uma aplicação em Streamlit criada para consultar cotações de frete na API da Intelipost a partir de um CEP de origem, uma lista de CEPs de destino e as dimensões do produto. A interface também permite importar CEPs por planilha Excel, usar uma base padrão de CEPs, baixar resultados em Excel e acompanhar logs técnicos da execução.

O objetivo principal é facilitar análises comparativas de frete, exibindo tanto os detalhes de cada opção retornada quanto indicadores consolidados, como menor frete, maior frete, frete médio, menor prazo, maior prazo e prazo médio.

Principais recursos
Cotação de frete para múltiplos CEPs em uma única execução.

Entrada manual de CEPs de destino.

Importação de CEPs por arquivo Excel, lendo os valores da coluna A.

Inclusão opcional de uma base padrão de CEPs a partir do arquivo ceps_padrao.xlsx.

Escolha entre modo de tabelas vigentes e tabelas em rascunho.

Exibição de métricas consolidadas em formato de big numbers.

Visualização de médias por transportadora.

Visualização de médias por CEP de destino.

Exportação dos resultados, logs e erros em arquivos Excel.

Área de debug com payload e headers da última requisição.

Como a ferramenta funciona
A aplicação monta um payload com os dados de origem, destino, dimensões, peso e valor do produto e envia esse conteúdo para o endpoint quote_by_product da Intelipost. Para cada CEP de destino processado, a resposta da API é analisada e as opções de entrega retornadas são convertidas em uma estrutura tabular com transportadora, prazo e valor do frete.

Depois disso, os dados são organizados em DataFrames do pandas para exibição na interface, cálculo das métricas e geração dos arquivos Excel para download. Quando a base padrão de CEPs está disponível, a aplicação também pode enriquecer a saída com informações como estado, cidade e tipo.

Requisitos
Antes de usar a ferramenta, confirme os itens abaixo:

Python 3.10 ou superior.

Bibliotecas instaladas: streamlit, pandas, requests, openpyxl.

Uma API Key válida da Intelipost.

Arquivo ceps_padrao.xlsx disponível no mesmo diretório do script, caso deseje usar a base padrão.

Exemplo de instalação das dependências:

bash
pip install streamlit pandas requests openpyxl
Estrutura esperada dos arquivos
A ferramenta pode operar apenas com digitação manual dos CEPs, mas alguns recursos dependem de arquivos auxiliares.

Arquivo de CEPs padrão
O arquivo ceps_padrao.xlsx deve conter colunas equivalentes a:

estado ou uf

cidade ou municipio

cep

tipo

A aplicação faz uma normalização automática dos nomes das colunas e também converte nomes completos de estados para suas siglas.

Arquivo Excel importado pelo usuário
Ao importar uma planilha de CEPs pela interface, a leitura é feita usando apenas a coluna A. Cada linha válida da coluna é considerada um CEP de destino.

Como executar
No terminal, acesse a pasta do projeto e execute:

bash
streamlit run nome_do_arquivo.py
Depois disso, o Streamlit abrirá a aplicação no navegador.

Como usar a interface
1. Informar a API Key
Na barra lateral, preencha o campo API Key Intelipost. Essa chave fica armazenada apenas na sessão atual do navegador e não é salva em arquivo.

2. Escolher o modo de cotação
Na barra lateral, selecione uma das opções:

Tabelas vigentes: envia a requisição padrão.

Tabelas em rascunho: adiciona os headers debug e logistic-contract-mode=DRAFT.

Use o modo em rascunho quando for necessário testar contratos logísticos ainda não publicados.

3. Informar o CEP de origem
No campo CEP origem, digite o CEP de saída da mercadoria. Esse valor será usado em todas as cotações da execução.

4. Informar os CEPs de destino
A ferramenta aceita CEPs de destino de três formas:

Digitação manual no campo de texto.

Importação por planilha Excel.

Inclusão automática da base padrão.

Essas fontes podem ser combinadas. A aplicação remove duplicidades antes de iniciar o processamento.

5. Definir peso, dimensões e valor
Preencha os campos:

Peso em kg.

Largura em cm.

Altura em cm.

Comprimento em cm.

Valor do item em reais.

A ferramenta converte automaticamente números em formato brasileiro, como 100,00.

6. Executar a cotação
Clique no botão Cotar fretes. A aplicação irá processar os CEPs um a um, mostrar o progresso na tela e registrar tanto os sucessos quanto os erros ocorridos.

7. Limpar resultados
O botão Limpar resultados remove da sessão os dados carregados da última execução, incluindo opções retornadas, logs e erros.

Indicadores exibidos
Após uma execução com retorno de dados, a aplicação mostra métricas consolidadas em destaque. Os big numbers disponíveis são:

CEPs cotados.

Opções retornadas.

Menor frete.

Maior frete.

Frete médio.

Menor prazo.

Maior prazo.

Prazo médio.

Esses indicadores ajudam a identificar rapidamente extremos e comportamento médio das cotações retornadas.

Abas da aplicação
Cotações
Exibe todas as opções de frete retornadas pela API, normalmente com colunas como estado, cidade, tipo, destino, transportadora, prazo e valor do frete. Também oferece download em Excel.

Médias por transportadora
Agrupa os resultados por transportadora e calcula:

Média de valor de frete.

Média de prazo.

Quantidade de opções retornadas.

Essa aba ajuda a comparar o comportamento médio de cada transportadora.

Médias por CEP
Agrupa os resultados por CEP de destino e calcula:

Média de valor de frete.

Média de prazo.

Quantidade de opções retornadas.

Se a base padrão estiver disponível, a visualização pode incluir também estado, cidade e tipo.

Execução
Mostra o resumo operacional do processamento de cada CEP, incluindo status, quantidade de opções e modo de tabela utilizado. Essa aba também pode ser exportada em Excel.

CEPs com erro
Lista os CEPs que apresentaram erro HTTP, exceção ou ausência de opções retornadas. Essa aba é útil para investigação e reprocessamento.

Tratamento de erros
A aplicação registra diferentes tipos de problema durante o processamento:

SEM_OPCOES: a API respondeu com sucesso, mas não retornou opções de entrega.

HTTP: houve erro de resposta HTTP.

EXCEPTION: ocorreu uma exceção no processamento local ou na chamada.

A mensagem detalhada do erro é armazenada para análise posterior e pode ser baixada em planilha.

Área de debug
Na seção Debug técnico, a ferramenta exibe:

Último payload enviado.

Últimos headers enviados.

Esse recurso é útil para homologação, troubleshooting e validação do conteúdo transmitido à API.

Regras de negócio implementadas
A ferramenta adota algumas regras importantes:

Os CEPs são unificados antes da execução para evitar duplicidade.

A API Key permanece somente na sessão atual.

O modo rascunho altera os headers da requisição.

A planilha de CEPs padrão passa por validação de colunas obrigatórias.

UFs são normalizadas tanto por sigla quanto por nome completo.

As respostas da API são convertidas para estrutura tabular com foco em análise.

Exemplo de uso
Um cenário comum é o seguinte:

Informar a API Key.

Definir o CEP de origem.

Carregar uma lista de CEPs de clientes por Excel.

Adicionar também os CEPs padrão.

Preencher peso, dimensões e valor do produto.

Executar a cotação.

Avaliar os indicadores de menor, maior e média de frete e prazo.

Exportar os resultados para análise externa.

Boas práticas
Validar se a API Key está correta antes de processar muitos CEPs.

Conferir se os CEPs da planilha estão realmente na coluna A.

Garantir que o arquivo ceps_padrao.xlsx tenha as colunas esperadas.

Usar o modo rascunho apenas quando necessário.

Exportar logs e erros após execuções grandes para rastreabilidade.

Revisar a aba de debug quando houver comportamento inesperado na API.

Possíveis evoluções
A ferramenta pode ser expandida com melhorias como:

Filtros por transportadora.

Gráficos comparativos de frete e prazo.

Paralelização das requisições.

Persistência opcional de histórico de execuções.

Validação visual de CEPs inválidos antes da chamada.

Indicadores segmentados por estado, cidade ou tipo.

Conclusão
A aplicação oferece uma forma prática de consultar, consolidar e analisar cotações de frete da Intelipost em lote, com foco em produtividade operacional e leitura gerencial. Com poucos passos, é possível comparar transportadoras, identificar extremos de custo e prazo, exportar os dados e investigar erros de forma estruturada.