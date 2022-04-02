###############################################################################
# 1. CARREGAMENTO DOS PACOTES
###############################################################################

library(dplyr)
library(arules)
library(arulesViz)
library(htmlwidgets)
library(writexl)
library(openxlsx)
options(warn=-1)

###############################################################################
# 2. CARREGAMENTO DOS DADOS
###############################################################################

df <- read.csv('dataset_bd3.csv')

###############################################################################
# 3. EXPLORAÇÃO E FORMATAÇÃO DO DATASET
###############################################################################

View(df)
dim(df)
summary(df)
str(df)

# O Dataset possui linhas em branco (linhas ímpares), vamos separá-las
linhas_pares <- seq(2, nrow(df), 2)

# Criação de um novo df com apenas as linhas pares
df <- df[linhas_pares, ]
View(df)

# Verifica se há valores ausentes no primeito item da compra (Primeira Coluna)
which(nchar(trimws(df$Item01))==0)

# Número de itens distintos
n_distinct(df)

# Selecionaremos apenas as linhas que contenham mais de um produto
# Expressão retorna TRUE e FALSE para valores "em branco"
df <- df[!grepl('^\\s*$', df$Item02), ]
View(df)

# Número de itens distintos
n_distinct(df)

# Conversão das variáveis para o tipo fator (Selecionando Apenas 6 Colunas)
for(i in 1:6){
  df[,i] <- as.factor(df[,i])
}

# Verificação do Dataset
str(df)
summary(df)
View(df)

# Criação de um novo DataFrame com Apenas 6 colunas
df_split <- split(df$Item01, df$Item02, df$Item03,
                  df$Item04, df$Item05, df$Item06,)
View(df_split)

# Criação e Exportação dos valores únicos da primeira coluna do DataFrame
dfUnicos <- as.data.frame(unique(df[, 1]))
write_xlsx(dfUnicos, 'dfUnicos.xlsx')
itemFrequencyPlot(transacoes, top=20, horiz=T) + 
  title('TOP 20 Produtos Mais Frequentes')
wb <- loadWorkbook('dfUnicos.xlsx')
insertPlot(wb, 1, width = 15, height = 10, 
           fileType = "png", units = "in", startCol = 5)
saveWorkbook(wb, 'dfUnicos.xlsx', overwrite = TRUE)

itemFrequencyPlot(transacoes, top=20, horiz=T, title='sadsa') + title('sadasdsd')
?itemFrequencyPlot

?colours
###############################################################################
# 4. EXTRAÇÃO DAS REGRAS DE ASSOCIAÇÃO
###############################################################################

# Transações - Convertendo o DataFrame no Formato transactions
transacoes <- as(df_split, 'transactions')

# Criação da função que irá gerar as regras para um determinado produto
gerarRegra <- function(nomeProduto){
  regra <- apriori(transacoes,
          parameter=list(conf=0.5, 
                         minlen=3,
                         supp= 0.2,
                         target='rules'),
          appearance=list(rhs=nomeProduto,
                          default='lhs'))
  regra <- regra[!is.redundant(regra)]}

# Criação da função que irá inspecionar a regra criada
inspecionarRegra <- function(nomeRegra, metrica){
  inspect(head(sort(nomeRegra, by=metrica, decreasing=T), 1))} 

# Criação da função para plotagem do gráfico
plotGrafico <- function(nomeRegra){
  plot(nomeRegra,
       measure='support',
       shading='confidence',
       method='graph',
       engine='html')}

###############################################################################
# 5. APLICAÇÃO DO MODELO (TESTE)
###############################################################################

# Produto 1: Dust-Off Compressed Gas 2 pack
regra_produto1 <- gerarRegra('Dust-Off Compressed Gas 2 pack')
inspecionarRegra(regra_produto1, 'support')
plotGrafico(regra_produto1)

# Produto 2: HP 61 ink
regra_produto2 <- gerarRegra('HP 61 ink')
inspecionarRegra(regra_produto2, 'confidence')
plotGrafico(regra_produto2)

# Produto 3: VIVO Dual LCD Monitor Desk mount
regra_produto3 <- gerarRegra('VIVO Dual LCD Monitor Desk mount')
inspecionarRegra(regra_produto3, 'confidence')
plotGrafico(regra_produto3)

###############################################################################
# 6. SALVANDO E EXPORTANTO O RESULTADO
###############################################################################

# Criação da função que irá retornar um arquivo em .XLSX com as regras
exportarResultado <- function(nomeProduto){
  regraProduto <- gerarRegra(nomeProduto)
  dfProduto <- as(regraProduto, 'data.frame')
  path_name <- paste(nomeProduto, '.xlsx', sep='')
  write_xlsx(dfProduto, path_name)
  grafico <- plot(regraProduto, method = 'grouped')
  wb <- loadWorkbook(path_name)
  print(grafico)
  insertPlot(wb, 1, width = 10, height = 7, 
             fileType = "png", units = "in", startCol = 8)
  saveWorkbook(wb, path_name, overwrite = TRUE)
}

###############################################################################
# 7. TESTANDO O RESULTADO COM PRODUTOS DIFERENTES
###############################################################################
exportarResultado('Dust-Off Compressed Gas 2 pack')
exportarResultado('Apple Lightning to Digital AV Adapter')
exportarResultado('HP 61 ink')
exportarResultado('Nylon Braided Lightning to USB cable')
exportarResultado('VIVO Dual LCD Monitor Desk mount')
