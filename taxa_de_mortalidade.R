# instalação do Pacote
install.packages("readxl")

# Bibliotecas Usadas
# Biblioteca usada para ler planilhas em Excel
library(readxl)

# Biblioteca usada para chamar a função "Describe" para analisar Medidas de Posição
library(psych)

# Usado para extrair Gráficos
library(lessR)

# Escolher local das Tabelas
file.choose()

# Nomear caminho da Tabela
taxa_mortalidade = "D:\\Faculdade Mackenzie\\2º Semestre\\07 - Projeto Aplicado\\Aula 2\\Projeto\\tabelas\\Taxa de Mortalidade.xlsx"

# Verificar quais tabelas existem na Planilha
excel_sheets(taxa_mortalidade)

# Carregar tabelas - Dataset
npessoas_doenca = read_excel(taxa_mortalidade, sheet = "Numero Pessoas - Doença",
                             range = "A3:Q11")
obito_doenca = read_excel(taxa_mortalidade, sheet = "Óbito - Doença",
                          range = "A3:Q11")

# Usando a função Head para verificar as informações da Tabela que temos
head(npessoas_doenca)
head(obito_doenca)
# Verificando os Tipos de Dados que temos na Tabela
str(npessoas_doenca)
str(obito_doenca)
                                                    # Medidas de Posição

# Usando "Summary" conseguimos listar: Minimo, 1º Quartil, Mediana, Média, 3º Quartil e o Máximo em apenas um comando
summary(npessoas_doenca)
summary(obito_doenca)

# Usando "Describe" da Biblioteca "psych", para extração de algus dados, como: Média, Mediana, Desvio Padrão,
#Mínimo, Máximo...
psych::describe(npessoas_doenca)
psych::describe(obito_doenca)

# "Média" de mortalidade atribuída a doenças cardiovasculares, câncer, diabetes ou doenças respiratórias crônicas 
# (Idade de 30 a 69) - Homens, Mulheres e o Total, separados por cada Ano no Brasil (2015 a 2019)
# Função "mean" foi usada para ver uma Média de maneira individual


mean(npessoas_doenca$Homens_2015)
mean(npessoas_doenca$Homens_2016)
mean(npessoas_doenca$Homens_2017)
mean(npessoas_doenca$Homens_2018)
mean(npessoas_doenca$Homens_2019)

mean(npessoas_doenca$Mulheres_2015)
mean(npessoas_doenca$Mulheres_2016)
mean(npessoas_doenca$Mulheres_2017)
mean(npessoas_doenca$Mulheres_2018)
mean(npessoas_doenca$Mulheres_2019)

mean(npessoas_doenca$Total_2015)
mean(npessoas_doenca$Total_2016)
mean(npessoas_doenca$Total_2017)
mean(npessoas_doenca$Total_2018)
mean(npessoas_doenca$Total_2019)

# óbito

mean(obito_doenca$Homens_2015)
mean(obito_doenca$Homens_2016)
mean(obito_doenca$Homens_2017)
mean(obito_doenca$Homens_2018)
mean(obito_doenca$Homens_2019)

mean(obito_doenca$Mulheres_2015)
mean(obito_doenca$Mulheres_2016)
mean(obito_doenca$Mulheres_2017)
mean(obito_doenca$Mulheres_2018)
mean(obito_doenca$Mulheres_2019)

mean(obito_doenca$Total_2015)
mean(obito_doenca$Total_2016)
mean(obito_doenca$Total_2017)
mean(obito_doenca$Total_2018)
mean(obito_doenca$Total_2019)

# "Mediana" de mortalidade atribuída a doenças cardiovasculares, câncer, diabetes ou doenças respiratórias crônicas 
# (Idade de 30 a 69) - Homens, Mulheres e o Total, separados por cada Ano no Brasil (2015 a 2019)
# Função "mean" foi usada para ver uma Média de maneira individual

median(npessoas_doenca$Homens_2015)
median(npessoas_doenca$Homens_2016)
median(npessoas_doenca$Homens_2017)
median(npessoas_doenca$Homens_2018)
median(npessoas_doenca$Homens_2019)

median(npessoas_doenca$Mulheres_2015)
median(npessoas_doenca$Mulheres_2016)
median(npessoas_doenca$Mulheres_2017)
median(npessoas_doenca$Mulheres_2018)
median(npessoas_doenca$Mulheres_2019)

median(npessoas_doenca$Total_2015)
median(npessoas_doenca$Total_2016)
median(npessoas_doenca$Total_2017)
median(npessoas_doenca$Total_2018)
median(npessoas_doenca$Total_2019)


# óbito


median(obito_doenca$Homens_2015)
median(obito_doenca$Homens_2016)
median(obito_doenca$Homens_2017)
median(obito_doenca$Homens_2018)
median(obito_doenca$Homens_2019)

median(obito_doenca$Mulheres_2015)
median(obito_doenca$Mulheres_2016)
median(obito_doenca$Mulheres_2017)
median(obito_doenca$Mulheres_2018)
median(obito_doenca$Mulheres_2019)

median(obito_doenca$Total_2015)
median(obito_doenca$Total_2016)
median(obito_doenca$Total_2017)
median(obito_doenca$Total_2018)
median(obito_doenca$Total_2019)


# Não tem como Efetuar a "Moda", pois os dados usados não tem uma "Frequência" em sua base

# Usando a função "quantile" para analisar o quartil de cada item da tabela (Homens, Mulheres e Total)

quantile(npessoas_doenca$Homens_2015)
quantile(npessoas_doenca$Homens_2016)
quantile(npessoas_doenca$Homens_2017)
quantile(npessoas_doenca$Homens_2018)
quantile(npessoas_doenca$Homens_2019)

quantile(npessoas_doenca$Mulheres_2015)
quantile(npessoas_doenca$Mulheres_2016)
quantile(npessoas_doenca$Mulheres_2017)
quantile(npessoas_doenca$Mulheres_2018)
quantile(npessoas_doenca$Mulheres_2019)

quantile(npessoas_doenca$Total_2015)
quantile(npessoas_doenca$Total_2016)
quantile(npessoas_doenca$Total_2017)
quantile(npessoas_doenca$Total_2018)
quantile(npessoas_doenca$Total_2019)


# óbito


quantile(obito_doenca$Homens_2015)
quantile(obito_doenca$Homens_2016)
quantile(obito_doenca$Homens_2017)
quantile(obito_doenca$Homens_2018)
quantile(obito_doenca$Homens_2019)

quantile(obito_doenca$Mulheres_2015)
quantile(obito_doenca$Mulheres_2016)
quantile(obito_doenca$Mulheres_2017)
quantile(obito_doenca$Mulheres_2018)
quantile(obito_doenca$Mulheres_2019)

quantile(obito_doenca$Total_2015)
quantile(obito_doenca$Total_2016)
quantile(obito_doenca$Total_2017)
quantile(obito_doenca$Total_2018)
quantile(obito_doenca$Total_2019)

                                                    # Medidas de Dispersão

# Usando a Função "var" para calcular a Variância entre os dados de maneira Individual

var(npessoas_doenca$Homens_2015)
var(npessoas_doenca$Homens_2016)
var(npessoas_doenca$Homens_2017)
var(npessoas_doenca$Homens_2018)
var(npessoas_doenca$Homens_2019)

var(npessoas_doenca$Mulheres_2015)
var(npessoas_doenca$Mulheres_2016)
var(npessoas_doenca$Mulheres_2017)
var(npessoas_doenca$Mulheres_2018)
var(npessoas_doenca$Mulheres_2019)

var(npessoas_doenca$Total_2015)
var(npessoas_doenca$Total_2016)
var(npessoas_doenca$Total_2017)
var(npessoas_doenca$Total_2018)
var(npessoas_doenca$Total_2019)

# Óbito

var(obito_doenca$Homens_2015)
var(obito_doenca$Homens_2016)
var(obito_doenca$Homens_2017)
var(obito_doenca$Homens_2018)
var(obito_doenca$Homens_2019)

var(obito_doenca$Mulheres_2015)
var(obito_doenca$Mulheres_2016)
var(obito_doenca$Mulheres_2017)
var(obito_doenca$Mulheres_2018)
var(obito_doenca$Mulheres_2019)

var(obito_doenca$Total_2015)
var(obito_doenca$Total_2016)
var(obito_doenca$Total_2017)
var(obito_doenca$Total_2018)
var(obito_doenca$Total_2019)

# Usando a Função "sd" para encontrar o Desvio Padrão de maneira individual
sd(npessoas_doenca$Homens_2015)
sd(npessoas_doenca$Homens_2016)
sd(npessoas_doenca$Homens_2017)
sd(npessoas_doenca$Homens_2018)
sd(npessoas_doenca$Homens_2019)

sd(npessoas_doenca$Mulheres_2015)
sd(npessoas_doenca$Mulheres_2016)
sd(npessoas_doenca$Mulheres_2017)
sd(npessoas_doenca$Mulheres_2018)
sd(npessoas_doenca$Mulheres_2019)

sd(npessoas_doenca$Total_2015)
sd(npessoas_doenca$Total_2016)
sd(npessoas_doenca$Total_2017)
sd(npessoas_doenca$Total_2018)
sd(npessoas_doenca$Total_2019)

# Óbito

sd(obito_doenca$Homens_2015)
sd(obito_doenca$Homens_2016)
sd(obito_doenca$Homens_2017)
sd(obito_doenca$Homens_2018)
sd(obito_doenca$Homens_2019)

sd(obito_doenca$Mulheres_2015)
sd(obito_doenca$Mulheres_2016)
sd(obito_doenca$Mulheres_2017)
sd(obito_doenca$Mulheres_2018)
sd(obito_doenca$Mulheres_2019)

sd(obito_doenca$Total_2015)
sd(obito_doenca$Total_2016)
sd(obito_doenca$Total_2017)
sd(obito_doenca$Total_2018)
sd(obito_doenca$Total_2019)