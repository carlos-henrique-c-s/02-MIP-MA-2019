
################################################################################
#                                                                              #
#                 ANALISE INSUMO-PRODUTO: APLICACAO MARANHAO                   #
#                                                                              #
################################################################################


###
### Iniciando/limpando objetos (1)
###

rm(list = ls())


###
### Definindo opcao de codificacao dos caracteres e linguagem (2)
###

aviso = getOption("warn")
options(warn=-1)
options(encoding="latin1")
options(warn=aviso)
rm(aviso)


###
### DEFININDO OPCAO DE EXIBICAO DE NUMEROS SEM EXPONENCIAL (3)
###

aviso <- getOption("warn")
options(warn=-1)
options(scipen=999)
options(warn=aviso)
rm(aviso)


###
### Pacotes necessários - instalação (4)
###

# Carregar planilhas

install.packages('openxlsx')

# Criacao de tabelas personalizadas

install.packages('flextable')

# Criacao de relatorios dinamicos

install.packages('knitr')

# Criacao de tabelas complexas

install.packages('kableExtra')

# Manipulacao e tratamento de dados

install.packages('dplyr')

# Visualizacao grafica de dados

install.packages('ggplot2')

# Funções de escala para visualização

install.packages('scales')

# Fornece geoms de texto e rótulo para 'ggplot2' que ajudam a evitar rótulos de texto sobrepostos

install.packages('ggrepel')

# Versão moderna dos quadros de dados

install.packages('tibble')

# Fornece várias funções de nível de usuário para trabalhar com gráficos de "grade"

install.packages('gridExtra')

# Pacote para adicionar imagens em gráficos

install.packages("ggimage")

# Pacote para adicionar GRIDES DE gráficos

install.packages("patchwork")


###
### Pacotes necessários - carregamento (4)
###

library('openxlsx')
library('flextable')
library('knitr')
library('kableExtra')
library('dplyr')
library('ggplot2')
library('scales')
library('ggrepel')
library('tibble')
library('gridExtra')
library("ggimage")
library("devtools")
library("patchwork")


###
### Base de dados- Extracao - carregamento dos data.frames (5)
###

# Apos baixar a planilha "MIP_MARANHÃO" save-a e defina o diretorio com a função
# "setwd" na mesma pasta de arquivos

## Consumo Intermediario (CI) - (5.1)

CI = round(read.xlsx("MIP_MARANHÃO.xlsx", sheet = "CI", colNames = FALSE))

## Demanda Final (DF) - (5.2)

DF = round(read.xlsx("MIP_MARANHÃO.xlsx", sheet = "DF", colNames = FALSE))

## Valor Bruto da Producao (VBP) - (5.3)

VBP = round(read.xlsx("MIP_MARANHÃO.xlsx", sheet = "VBP", colNames = FALSE))

## VAlor Adicionado (VA) - (5.4)

VAB = round(read.xlsx("MIP_MARANHÃO.xlsx", sheet = "VAB", colNames = FALSE))

## Remuneracoes (REM) - (5.5)

REM = round(read.xlsx("MIP_MARANHÃO.xlsx", sheet = "REM", colNames = FALSE))

## Pessoal Ocupado (PO) - (5.6)

PO = round(read.xlsx("MIP_MARANHÃO.xlsx", sheet = "PO", colNames = FALSE))

## Consumo das Familias (CF) - (5.7)

CF = round(read.xlsx("MIP_MARANHÃO.xlsx", sheet = "CF", colNames = FALSE))

## Setor de Pagamentos (SP) - (5.8)

SP = round(read.xlsx("MIP_MARANHÃO.xlsx", sheet = "SP", colNames = FALSE))

## Setores (SET) - (5.9)

SET = read.xlsx("MIP_MARANHÃO.xlsx", sheet = "SET", colNames = FALSE)

colnames(SET) = "Setores"

## Impostos (T) - (5.9)

TRIB = round(read.xlsx("MIP_MARANHÃO.xlsx", sheet = "TRIB", colNames = FALSE))


###
### Base de dados- Manipulacao - transformando os dfs em matrizes/vetores (6)
###

CI = data.matrix(CI) # Consumo intermediario
DF = data.matrix(DF) # Demanda final
VBP = data.matrix(VBP) # Valor bruto da producao
VBP = as.vector(VBP) # Valor bruto da producao
VAB = data.matrix(VAB) # Valor adicionado
REM = data.matrix(REM) # Remuneracoes
PO = data.matrix(PO) # Pessoal ocupado
CF = data.matrix(CF) # Consumo das familias
SP = data.matrix(SP) # Setor de pagamentos
TRIB = data.matrix(TRIB) # Impostos


###
### Modelo de Insumo-Produto aberto (7)
###

# Produto é igual à inversa de Leontief vezes a demanda final
#
#       VBP = DF * (I - A)^(-1) = DF * B
#
# onde A = coeficientes tecnicos diretos/razao insumo-produto
# onde B = (I - A)^(-1) e a matriz inversa de Leontief

## Matriz de coeficientes tecnicos diretos/razao insumo-produto (7.1)

A = round(CI %*% diag(1 / VBP), digits = 4)

## Matriz de coeficientes tecnicos diretos e indiretos/inversa de Leontief (7.2)
# construcao da matriz identidade 36 x 36

n = length(VBP)
I = diag(n)

## A matriz inversa de Leontief (7.3)

B = round(solve(I - A), digits = 4)


###
### Modelo de Insumo-Produto fechado (8)
###

## hc = coeficientes de consumo (8.1)

CF = data.matrix(CF)

hc = round((CF / sum(REM)), digits = 4)

## hr = coeficientes de remuneração do trabalho (8.2)

hr = round(REM / VBP, digits = 4)

hr = t(hr)

## Matriz de coeficientes tecnicos diretos/razao insumo-produto - modelo fechado (8.3)

A_fechado = matrix(NA, ncol = n + 1, nrow = n + 1)

A_fechado = round(rbind(cbind(A, hc), cbind(hr, 0)), digits = 4)

rm(hc, hr)

IF = diag(n + 1)

## A matriz inversa de Leontief - modelo fechado (8.4)

B_fechado = round(solve(IF - A_fechado), digits = 4)


###
### Modelo pelo lado da oferta (9)
###

## Matriz de coeficientes técnicos pelo lado da oferta (9.1)

OF = round((diag(1 / VBP) %*% CI), digits = 4)

## Matriz inversa de Grosh (9.2)

G = round(solve(I - OF), digits = 4)


###
### MULTIPLICADORES DE PRODUCAO (10)
###

## Multiplicador de Producao Simples - MPS (10.1)

MPS = round(colSums(B), digits = 4)

# Decomposicao 1: Intra

intra1 = B[1:18, 1:18]

intra2 = B[19:36, 19:36]

mult_intra1 = round(colSums(intra1), digits = 4)

rm(intra1)

mult_intra2 = round(colSums(intra2), digits = 4)

rm(intra2)

Intra = c(mult_intra1, mult_intra2)

rm(mult_intra1, mult_intra2)

## Multiplicador de Produção Total - MPT (10.2)

MPT = round(colSums(B_fechado[ , 1:n]), digits = 4)

## Multiplicador de Produção Total Truncado - MPTT (10.3)

MPTT = round(colSums(B_fechado[1:n, 1:n]), digits = 4)

## Multiplicadores (10.4)

MultProd = cbind(SET, MPS, Intra, MPT, MPTT) %>% 
  as.data.frame()

rm(Intra, MPS, MPT, MPTT)

MultProd = MultProd %>% 
  mutate("Setores" = X1,
         "Inter" = round(MPS - Intra, digits = 4),
         "Direto" = round(colSums(A), digits = 4),
         "Indireto" = round(MPS - Direto, digits = 4),
         "MPT (EI)" = MPT - MPS,
         "MPTT (EI)" = MPTT - MPS) %>% 
  select(Setores, MPS, Intra, Inter, Direto, Indireto, MPT, `MPT (EI)`, MPTT, `MPTT (EI)`)

MultProd = MultProd %>% 
  select(Setores, MPS, Intra, Inter, Direto, Indireto, MPT, MPTT)

# Tabela 1 - Multiplicadores

t1 = MultProd[19:36, ] %>% 
  flextable() %>% 
  colformat_double(j = 2:8, decimal.mark = ",", digits = 4) %>% 
  autofit(add_w = 0, add_h = 0) %>% 
  theme_zebra() %>% 
  align(align = "center", part = "header", i = 1) %>% 
  #set_caption(caption = "Multiplicadores de produção simples, intra-regional, inter-regional, totais e truncados") %>% 
  add_footer_lines("Fonte: Departamento de Contas Regionais e Finanças Públicas (DECRE-IMESC) - 2019") %>% 
  flextable::footnote(i = 1, j = 2:10,
                      value = as_paragraph(c("Efeito total no modelo aberto (Efeito direto + indireto)", 
                                             "Efeito intersetorial direto no estado maranhense para atender à variação de demanda", 
                                             "Parte do efeito total da variação de demanda no estado que vaza para o resto do Brasil",
                                             "Coeficientes técnicos ou coeficientes diretos - relação insumo-produto",
                                             "Efeito resultante do encadeamento produtivo com outros setores",
                                             "Efeito Total no modelo fechado (Efeito Direto + Indireto + Induzido - endogeniza as famílias)",
                                             "Efeito decorrente das famílias dentro do modelo de insumo-produto",
                                             "Efeito dos NxN setores no modelo fechado",
                                             "Efeito decorrente das famílias dentro do modelo de insumo-produto - NxN setores.")),
           ref_symbols = c("a", "b", "c", "d", "e", "f", "g", "h", "i"),
           part = "header",
           inline = T,
           sep = "; ")

save_as_pptx("Multiplicadores de produção - RBr" = t1,
             path = "Anexo 1.Multiplicadores de produção.pptx")

rm(t1)

# Grafico 1 - Multiplicadores de producao nos modelos aberto e fechado

g2 = MultProd %>% 
  head(n = 18) %>% 
  ggplot(aes(x = MPS, y = reorder(Setores, MPS, size = MPS))) +
  theme_minimal() + 
  geom_col(fill = "deepskyblue4") + 
  labs(title = "Multiplicador de produção simples - modelo aberto",
       x = " ",
       y = " ") +
  theme(axis.text = element_text(color = "black", size = 15)) +
  geom_text(aes(label = round(MPS, digits = 4)), vjust = 0.5, hjust = 1, size = 5, color = "white") +
  scale_x_continuous(breaks = seq(0, 3, .5), limits = c(0, 2.5))

g3 = MultProd %>% 
  head(n = 18) %>% 
  ggplot(aes(x = MPT, y = reorder(Setores, MPT, size = MPT))) +
  theme_minimal() + 
  geom_col(fill = "deepskyblue4") + 
  labs(title = "Multiplicador de produção total - modelo fechado",
       x = " ",
       y = " ") +
  theme(axis.text = element_text(color = "black", size = 15)) +
  geom_text(aes(label = round(MPT, digits = 4)), vjust = 0.5, hjust = 1, size = 5, color = "white") +
  scale_x_continuous(breaks = seq(0, 10, 2), limits = c(0, 10))

g1 = g2 + g3

ggsave(filename = "Grafico 1. Multiplicadores de producao nos modelos aberto e fechado.png",
       plot = g1,
       width = 16,
       height = 9,
       units = "in",
       dpi = 600)

rm(g1, g2, g3)

## Decomposição do multipicador de produção simples (10.5)

Decomp_MPS = MultProd %>% 
  mutate("Intra (%)" = round(((Intra / MPS) * 100), digits = 2), 
         "Inter (%)" = round(((Inter / MPS) * 100), digits = 2),
         "Intra Líq. (%)" = round((((Intra - 1) / (MPS - 1)) * 100), digits = 2),
         "Inter Líq. (%)" = round(((Inter / (MPS - 1)) * 100), digits = 2)) %>% 
  select(-Intra, -Inter, -MPT, -MPTT, -Direto, -Indireto) %>% 
  mutate(`Intra Líq. (%)` = case_when(is.nan(`Intra Líq. (%)`) ~ 00.00,
                                    TRUE ~`Intra Líq. (%)`),
         `Inter Líq. (%)` = case_when(is.nan(`Inter Líq. (%)`) ~ 00.00,
                                    TRUE ~ `Inter Líq. (%)`))

# Tabela 2 - Decomposição do multipicador de produção simples

t2 = Decomp_MPS[1:18, ] %>% 
  flextable() %>% 
  theme_zebra() %>% 
  add_header_row(colwidths = c(1, 3, 2),
                 values = c("", "Multiplicador de produção simples", "Multiplicador líquido simples")) %>% 
  align(align = "left", part = "header", i = 1) %>% 
  autofit(add_w = 0, add_h = 0) %>% 
  set_caption(caption = "Decomposição do multipicador de produção simples - Modelo aberto") %>% 
  dd_footer_lines("Departamento de Contas Regionais e Finanças Públicas (DECRE-IMESC) - 2019") %>% 
  flextable::footnote(i = 1, j = 2:3,
                      value = as_paragraph(c("O multiplicador de produção simples considera o impacto na produção provocado pela variação na demanda final, considerando a injeção inicial de uma unidade monetária",
                                             "O multiplicador de produção líquido dá o efeito multiplicador descontado da injeção inicial.")),
                      ref_symbols = c("a", "b"),
                      part = "header")

save_as_pptx("Decomposição do multipicador de produção simples - MA" = t2,
             path = "Tabela 2.Decomposição do multipicador de produção simples.pptx")

rm(t2)


# Grafico 2 - Decomposição do multipicador de produção simples

g2 = Decomp_MPS %>% 
  head(n = 18) %>% 
  ggplot(aes(x = MPS, y = reorder(Setores, MPS, size = MPS))) +
  geom_point(size = 3, color = "firebrick4") + 
  theme_minimal() + 
  labs(x = "Multiplicador de produção simples",
       y = "18 setores econômicos",
       title = "Multiplicador simples de produção - Modelo aberto",
       subtitle = "Maranhão",
       caption = "Departamento de Contas Regionais e Finanças Públicas (DECRE-IMESC) - 2022") + 
  theme(plot.title = element_text(hjust = .5),
        plot.subtitle = element_text(hjust = .5),
        plot.caption = element_text(hjust = 0),
        axis.text = element_text(size = 12),
        axis.text.x = element_text(size = 12),
        axis.text.y = element_text(size = 12)) + 
  scale_x_continuous(breaks = seq(1, 3, .2), limits = c(1, 2.4)) + 
  ggimage::geom_image(aes(x = 2.11, y = 1.5),
                      image = "imesc4.png",
                      size = .3)

ggsave(filename = "Gráfico 2. Multiplicador Simples de produção - modelo aberto.png",
       plot = g2,
       width = 12,
       height = 9,
       units = "in",
       dpi = 600)

rm(g2)

## Decomposição do multipicador de produção total - modelo fechado (10.6)

Decomp_MPT = MultProd %>% 
  mutate("MPS (%)" = round(((MPS / MPT) * 100), digits = 2), 
         "MPT-EI (%)" = round((((MPT - MPS) / MPT) * 100), digits = 2),
         "Direto (%)" = round(((Direto / MPT) * 100), digits = 2),
         "Indireto (%)" = round(((Indireto / MPT) * 100), digits = 2)) %>% 
  select(Setores, MPT, `MPS (%)`, `Direto (%)`, `Indireto (%)`, `MPT-EI (%)`)

# Tabela 3 - Decomposição do multipicador de produção total - modelo fechado

t3 = flextable(head(Decomp_MPT, n = 18)) %>% 
  align(align = "left", part = "header", i = 1) %>% 
  theme_zebra() %>% 
  autofit(add_w = 0, add_h = 0) %>% 
  set_caption(caption = "Decomposição do multipicador total de produção (MTP) - Modelo fechado") %>% 
  add_footer_lines("Departamento de Contas Regionais e Finanças Públicas (DECRE-IMESC) - 2022.") %>% 
  flextable::footnote(i = 1, j = 3:6,
                      value = as_paragraph(c("Participação do MPS (Efeito direto + indireto) no MTP;",
                                             "Percentual do efeito da introdução da famílias no modelo fechado;",
                                             "Percentual dos coeficientes técnicos/diretos no MTP;",
                                             "Participação do efeito intersetorial no MTP.")),
                      ref_symbols = c("  a", "  b", "  c", "  d"),
                      part = "header")

save_as_pptx("Decomposição do multipicador de produção total - modelo fechado (MA)" = t3,
             path = "Tabela 3.Decomposição do multipicador de produção total - modelo fechado.pptx")

rm(t3)


# Grafico 3 - Multipicador de produção total - modelo fechado

g3 = Decomp_MPT %>% 
  head(n = 18) %>% 
  ggplot(aes(x = MPT, y = reorder(Setores, MPT, size = MPT))) +
  geom_point(size = 3, color = "firebrick4") + 
  theme_minimal() + 
  labs(x = "Multiplicador de produção total",
       y = "18 setores econômicos",
       title = "Multiplicador de produção total - Modelo fechado",
       subtitle = "Maranhão",
       caption = "Departamento de Contas Regionais e Finanças Públicas (DECRE-IMESC) - 2022") + 
  theme(plot.title = element_text(hjust = .5),
        plot.subtitle = element_text(hjust = .5),
        plot.caption = element_text(hjust = 0),
        axis.text = element_text(size = 12),
        axis.text.x = element_text(size = 12),
        axis.text.y = element_text(size = 12)) + 
  scale_x_continuous(breaks = seq(1, 9, 1), limits = c(1, 9)) + 
  ggimage::geom_image(aes(x = 2.11, y = 1.5),
                      image = "imesc4.png",
                      size = .3)

ggsave(filename = "Gráfico 3. Multipicador de produção total - modelo fechado.png",
       plot = g3,
       width = 12,
       height = 9,
       units = "in",
       dpi = 600)

rm(g3)


## Decomposição do multipicador de produção total truncado - modelo fechado (10.7)

Decomp_MPTT = MultProd %>% 
  mutate("MPS (%)" = round(((MPS / MPTT) * 100), digits = 2), 
         "MPTT-EI (%)" = round((((MPTT - MPS) / MPTT) * 100), digits = 2),
         "Direto (%)" = round(((Direto / MPTT) * 100), digits = 2),
         "Indireto (%)" = round(((Indireto / MPTT) * 100), digits = 2)) %>% 
  select(Setores, MPTT, `MPS (%)`, `Direto (%)`, `Indireto (%)`, `MPTT-EI (%)`)

# Tabela 4 - Decomposição do multipicador de produção total truncado - modelo fechado

t4 = flextable(head(Decomp_MPTT, n = 18)) %>% 
  theme_zebra() %>% 
  align(align = "left", part = "header", i = 1) %>% 
  autofit(add_w = 0, add_h = 0) %>% 
  set_caption(caption = "Decomposição do multipicador de produção total truncado (MPTT) - Modelo fechado") %>% 
  add_footer_lines("Departamento de Contas Regionais e Finanças Públicas (DECRE-IMESC) - 2019.") %>% 
  flextable::footnote(i = 1, j = 3:6,
                      value = as_paragraph(c("Participação do MPS (Efeito direto + indireto) no MTPT;",
                                             "Percentual do efeito da introdução da famílias no modelo fechado-truncado;",
                                             "Percentual dos coeficientes técnicos/diretos no MTPT;",
                                             "Participação do efeito intersetorial no MTPT.")),
                      ref_symbols = c("  a", "  b", "  c", "  d"),
                      part = "header")

save_as_pptx("Decomposição do multipicador de produção total truncado - modelo fechado (MA)" = t4,
             path = "Tabela 4.Decomposição do multipicador de produção total truncado - modelo fechado.pptx")

rm(t4)


# Grafico 4 - Multipicador de produção total truncado - modelo fechado

g4 = Decomp_MPTT %>% 
  head(n = 18) %>% 
  ggplot(aes(x = MPTT, y = reorder(Setores, MPTT, size = MPTT))) +
  geom_point(size = 3, color = "firebrick4") + 
  theme_minimal() + 
  labs(x = "Multiplicador total de produção truncado",
       y = "18 setores econômicos",
       title = "Multipicador de produção total truncado - Modelo fechado",
       subtitle = "Maranhão",
       caption = "Departamento de Contas Regionais e Finanças Públicas (DECRE-IMESC) - 2022") + 
  theme(plot.title = element_text(hjust = .5),
        plot.subtitle = element_text(hjust = .5),
        plot.caption = element_text(hjust = 0),
        axis.text = element_text(size = 12),
        axis.text.x = element_text(size = 12),
        axis.text.y = element_text(size = 12)) + 
  scale_x_continuous(breaks = seq(1, 7, 1), limits = c(1, 7)) + 
  ggimage::geom_image(aes(x = 2.11, y = 1.5),
                      image = "imesc4.png",
                      size = .3)

ggsave(filename = "Gráfico 4. Multipicador de produção total truncado - modelo fechado.png",
       plot = g4,
       width = 12,
       height = 9,
       units = "in",
       dpi = 600)

rm(g4, Decomp_MPS, Decomp_MPT, Decomp_MPTT)


###
### MULTIPLICADORES DE EMPREGO (11)
###

## coeficientes de emprego - requisitos de emprego - CE (11.1)

CE = round((PO / VBP), digits = 4) %>% 
  as.vector()

matriz_CE = diag(CE)

## Matriz geradora de emprego - E (11.2)

E = round((matriz_CE %*% B), digits = 4)

## Multiplicador de Emprego Simples - MES (11.3)

MES = colSums(E)

# Decomposicao intra e inter

intra1 = E[1:18, 1:18] %>% 
  colSums()

intra2 = E[19:36, 19:36] %>% 
  colSums()

intra = c(intra1, intra2)

rm(intra1, intra2)

## Multiplicador de Emprego-Tipo I (MET1) - (11.4)

MET1 = round((MES / CE), digits = 4)

## Multiplicador de Emprego Total Truncado (METT) - (11.5)

ET = round((matriz_CE %*% B_fechado[1:n, 1:n]), digits = 4)

METT = colSums(ET)

## Multiplicador de Emprego-Tipo II (MET2) - (11.6)

MET2 = round((METT / CE), digits = 4)

## Multiplicadores (11.7)

MultEmp = data.frame("Setores" = SET$Setores,
                     "MES" = round(MES, digits = 2),
                     "Intra" = round(intra, digits = 2),
                     "Inter" = round(MES - intra, digits = 2),
                     "Direto" = round(CE, digits = 2),
                     "indireto" = round(MES - CE, digits = 2),
                     "METT" = round(METT, digits = 2),
                     "MET1" = round(MET1, digits = 2),
                     "MET2" = round(MET2, digits = 2))

rm(CE, matriz_CE, E, MES, intra, MET1, ET, METT, MET2)

## Tabela 5 - Multiplicadores de emprego (11.8)

t5 = MultEmp[1:18, ] %>% 
  flextable() %>% 
  theme_zebra() %>% 
  align(align = "center", part = "header", i = 1) %>% 
  colformat_double(j = c(2:9), big.mark = ",", decimal.mark = ",", digits = 2) %>% 
  autofit(add_w = 0, add_h = 0) %>% 
  set_caption(caption = "Multiplicadores de emprego simples e truncados") %>% 
  add_footer_lines("Fonte: Departamento de Contas Regionais e Finanças Públicas (DECRE-IMESC) - 2019. \nNota: MSE e MTET por milhão de reais. a interpretação segue: por exemplo, uma variação de demanda de R$ 1.000.000 no setor Agro, gera 7,8937 empregos na economia.") %>% 
  flextable::footnote(i = 1, j = 2:5,
                      value = as_paragraph(c("O Multiplicador Simples de Emprego (ou Gerador de Emprego) apresenta o impacto total (direto + indireto) em termos de empregos gerados na economia dada uma variação unitária da demanda final;", 
                                             "Multiplicador de Emprego Tipo 1: efeito multiplicador na economia para cada emprego gerado diretamente no setor específico;", 
                                             "O Multiplicador Total de Emprego Truncado mostra os impactos direto, indireto e induzido, em termos de empregos gerados na economia, dada uma variação unitária da demanda final;",
                                             "Multiplicador de Emprego Tipo 2: efeito multiplicador na economia, considerando o efeito induzido, dado cada emprego gerado no setor específico.")),
                      ref_symbols = c("a", "b", "c", "d"),
                      part = "header",
                      sep = "; ")

save_as_pptx("Multiplicadores de emprego - MA" = t5,
             path = "Tabela 5.Multiplicadores de emprego.pptx")

rm(t5)

## Grafico 5: Multiplicadores de emprego nos modelos aberto e fechado (11.9)

g = MultEmp %>% 
  head(n = 18) %>% 
  ggplot(aes(x = MES, y = reorder(Setores, MES, size = MES))) +
  theme_minimal() + 
  geom_col(fill = "deepskyblue4") + 
  labs(title = "Multiplicador simples de emprego - modelo aberto",
       x = " ",
       y = " ") +
  theme(axis.text = element_text(color = "black", size = 15)) +
  geom_text(aes(label = round(MES, digits = 0)), vjust = 0.5, hjust = 1, size = 5, color = "white") +
  scale_x_continuous(breaks = seq(0, 180, 20), limits = c(0, 180))

g1 = MultEmp %>% 
  head(n = 18) %>% 
  ggplot(aes(x = METT, y = reorder(Setores, METT, size = METT))) +
  theme_minimal() + 
  geom_col(fill = "deepskyblue4") + 
  labs(title = "Multiplicador total de emprego truncado - modelo fechado",
       x = " ",
       y = " ") +
  theme(axis.text = element_text(color = "black", size = 15)) +
  geom_text(aes(label = round(METT, digits = 0)), vjust = 0.5, hjust = 1, size = 5, color = "white") +
  scale_x_continuous(breaks = seq(0, 220, 20), limits = c(0, 220))

#

g5 = g + g1

ggsave(filename = "Grafico 5.Multiplicadores de emprego nos modelos aberto e fechado.png",
       plot = g5,
       width = 18,
       height = 10)

rm(g, g1, g5)


###
### Multiplicadores de Renda (12)
###

## Coeficientes de renda ou requisitos de renda (12.1)

CR = round((REM / VBP), digits = 4) %>% 
  as.vector()

matriz_CR = diag(CR)

## Matriz de coeficientes de renda (12.2)

Renda = round((matriz_CR %*% B), digits = 4)

## Multiplicador de Renda Simples ou Gerador de Renda - MRS (12.3)

MRS = round((colSums(Renda)), digits = 4)

# Decomposicao intra e inter

intra1 = Renda[1:18, 1:18] %>% 
  colSums()

intra2 = Renda[19:36, 19:36] %>% 
  colSums()

intra = c(intra1, intra2)

## Multiplicador de Renda - Tipo I (12.4)

MRT1 = round((MRS / CR), digits = 4)

## Multiplicador Total de Renda - truncado (12.5)

renda_trunc = round(matriz_CR %*% B_fechado[1:n, 1:n], digits = 4)

MRTT = round(colSums(renda_trunc), digits = 4)

## Multiplicador de Renda - Tipo II (12.6)

MRT2 = round(MRTT / CR, digits = 4)

## Multiplicadores de renda (12.7)

MultRend = data.frame("Setores" = SET$Setores,
                     "MRS" = round(MRS, digits = 4),
                     "Intra" = round(intra, digits = 4),
                     "Inter" = round(MRS - intra, digits = 4),
                     "Direto" = round(CR, digits = 4),
                     "indireto" = round(MRS - CR, digits = 4),
                     "MRTT" = round(MRTT, digits = 4),
                     "MRT1" = round(MRT1, digits = 4),
                     "MRT2" = round(MRT2, digits = 4))

rm(MRS, MRT1, MRTT, MRT2, CR, matriz_CR, Renda, renda_trunc, intra1, intra2, intra)

## Tabela 6 - Multiplicadores de renda (12.8)

t6 = MultRend[19:36, ] %>% 
  flextable() %>% 
  autofit(add_w = 0, add_h = 0) %>% 
  theme_zebra() %>% 
  align(align = "center", part = "header", i = 1) %>% 
  set_caption(caption = "Multiplicadores de renda simples e truncados") %>% 
  add_footer_lines("Fonte: Departamento de Contas Regionais e Finanças Públicas (DECRE-IMESC) - 2019. \nNota: MSR e MTRT por milhão de reais. A interpretação segue: por exemplo, uma variação de demanda de R$ 1.000.000 no setor Agro, gera 0,1977 de renda na economia.") %>% 
  flextable::footnote(i = 1, j = 2:5,
                      value = as_paragraph(c("O Multiplicador Simples de Renda (ou Gerador de Renda) apresenta o impacto total (direto + indireto) em termos de renda gerada na economia dada uma variação unitária da demanda final;", 
                                             "Multiplicador de Renda Tipo 1: efeito multiplicador na economia para cada unidade de renda gerada diretamente no setor específico;", 
                                             "O Multiplicador Total de Renda Truncado mostra os impactos direto, indireto e induzido, em termos de renda gerada na economia, dada uma variação unitária da demanda final;",
                                             "Multiplicador de Renda Tipo 2: efeito multiplicador na economia, considerando o efeito induzido, dada cada unidade de renda gerada no setor específico.")),
                      ref_symbols = c("a", "b", "c", "d"),
                      part = "header",
                      sep = "; ")

save_as_pptx("Multiplicadores de renda - Rbr" = t6,
             path = "Anexo 4.Multiplicadores de renda.pptx")

rm(t6)


## Grafico 6 - Multiplicadores de renda (12.9)

g = MultRend %>% 
  head(n = 18) %>% 
  ggplot(aes(x = MRT1, y = reorder(Setores, MRT1, size = MRT1))) +
  theme_minimal() + 
  geom_col(fill = "deepskyblue4") + 
  labs(title = "Multiplicador de renda tipo 1 - modelo aberto",
       x = " ",
       y = " ") +
  theme(axis.text = element_text(color = "black", size = 15)) +
  geom_text(aes(label = round(MRT1, digits = 4)), vjust = 0.5, hjust = 1, size = 5, color = "white") +
  scale_x_continuous(breaks = seq(0, 5, 1), limits = c(0, 5))


g1 = MultRend %>% 
  head(n = 18) %>% 
  ggplot(aes(x = MRT2, y = reorder(Setores, MRT2, size = MRT2))) +
  theme_minimal() + 
  geom_col(fill = "deepskyblue4") + 
  labs(title = "Multiplicador de renda tipo 2 - modelo fechado",
       x = " ",
       y = " ") +
  theme(axis.text = element_text(color = "black", size = 15)) +
  geom_text(aes(label = round(MRT2, digits = 4)), vjust = 0.5, hjust = 1, size = 5, color = "white") +
  scale_x_continuous(breaks = seq(0, 8, 1), limits = c(0, 8))

#

g6 = g + g1

ggsave(filename = "Grafico 6.Multiplicadores de renda.png",
       plot = g6,
       width = 18,
       height = 10)

rm(g, g1, g6)


###
### MULTIPLICADORES DO VALOR ADICIONADO (13)
###

## Coeficientes do Valor Adicionado Bruto - CVAB (13.1)

CVAB = round((VAB / VBP), digits = 4) %>% 
  as.vector()

matriz_CVAB = diag(CVAB)

## Matriz de coeficientes do VAB - VAB1 (13.2)

VAB1 = round((matriz_CVAB %*% B), digits = 4)

## Multiplicador do VAB Simples - ou Gerador de VAB - MVABS (13.3)

MVABS = round((colSums(VAB1)), digits = 4)

# Decomposicao intra e inter

intra1 = VAB1[1:18, 1:18] %>% 
  colSums()

intra2 = VAB1[19:36, 19:36] %>% 
  colSums()

intra = c(intra1, intra2)

## Multiplicador de VAB - Tipo I (13.4)

MVABT1 = round((MVABS / CVAB), digits = 4)

## Multiplicador Total de VAB - truncado (13.5)

VAB_trunc = round(matriz_CVAB %*% B_fechado[1:n, 1:n], digits = 4)

MVABTT = round(colSums(VAB_trunc), digits = 4)

## Multiplicador de VAB - Tipo II (13.6)

MVABT2 = round(MVABTT / CVAB, digits = 4)

## Multiplicadores de renda (13.7)

MultVAB = data.frame("Setores" = SET$Setores,
                      "MVABS" = round(MVABS, digits = 4),
                      "Intra" = round(intra, digits = 4),
                      "Inter" = round(MVABS - intra, digits = 4),
                      "Direto" = round(CVAB, digits = 4),
                      "indireto" = round(MVABS - CVAB, digits = 4),
                      "MVABTT" = round(MVABTT, digits = 4),
                      "MVABT1" = round(MVABT1, digits = 4),
                      "MVABT2" = round(MVABT2, digits = 4))

rm(CVAB, matriz_CVAB, VAB1, MVABS, intra1, intra2, intra, MVABT1, VAB_trunc, MVABTT, MVABT2)

## Tabela 7: Multiplicadores do VAB (13.8)

t7 = MultVAB[19:36, ] %>% 
  flextable() %>% 
  theme_zebra() %>% 
  align(align = "center", part = "header", i = 1) %>%
  autofit(add_w = 0, add_h = 0) %>% 
  set_caption(caption = "Multiplicadores de renda simples e truncados") %>% 
  add_footer_lines("Fonte: Departamento de Contas Regionais e Finanças Públicas (DECRE-IMESC) - 2019. \nNota: MSR e MTRT por milhão de reais. A interpretação segue: por exemplo, uma variação de demanda de R$ 1.000.000 no setor Agro, gera 0,1977 de renda na economia.") %>% 
  flextable::footnote(i = 1, j = 2:5,
                      value = as_paragraph(c("O Multiplicador Simples de Renda (ou Gerador de Renda) apresenta o impacto total (direto + indireto) em termos de renda gerada na economia dada uma variação unitária da demanda final;", 
                                             "Multiplicador de Renda Tipo 1: efeito multiplicador na economia para cada unidade de renda gerada diretamente no setor específico;", 
                                             "O Multiplicador Total de Renda Truncado mostra os impactos direto, indireto e induzido, em termos de renda gerada na economia, dada uma variação unitária da demanda final;",
                                             "Multiplicador de Renda Tipo 2: efeito multiplicador na economia, considerando o efeito induzido, dada cada unidade de renda gerada no setor específico.")),
                      ref_symbols = c("a", "b", "c", "d"),
                      part = "header",
                      sep = "; ")

save_as_pptx("Multiplicadores do VAB - RBr" = t7,
             path = "Anexo 5.Multiplicadores do VAB.pptx")

rm(t7)


## Grafico 7: Multiplicadores do VAB (13.9)

g = MultVAB %>% 
  head(n = 18) %>% 
  ggplot(aes(x = MVABT1, y = reorder(Setores, MVABT1, size = MVABT1))) +
  theme_minimal() + 
  geom_col(fill = "deepskyblue4") + 
  labs(title = "Multiplicadores do VAB tipo 1 - modelo aberto",
       x = " ",
       y = " ") +
  theme(axis.text = element_text(color = "black", size = 15)) +
  geom_text(aes(label = round(MVABT1, digits = 4)), vjust = 0.5, hjust = 1, size = 5, color = "white") +
  scale_x_continuous(breaks = seq(0, 4, 1), limits = c(0, 4))

g1 = MultVAB %>% 
  head(n = 18) %>% 
  ggplot(aes(x = MVABT2, y = reorder(Setores, MVABT2, size = MVABT2))) +
  theme_minimal() + 
  geom_col(fill = "deepskyblue4") + 
  labs(title = "Multiplicadores do VAB tipo 2 - modelo fechado",
       x = " ",
       y = " ") +
  theme(axis.text = element_text(color = "black", size = 15)) +
  geom_text(aes(label = round(MVABT2, digits = 4)), vjust = 0.5, hjust = 1, size = 5, color = "white") +
  scale_x_continuous(breaks = seq(0, 6, 1), limits = c(0, 6))

###

g7 = g + g1

ggsave(filename = "Grafico 7.Multiplicadores do VAB.png",
       plot = g7,
       width = 18,
       height = 10)

rm(g, g1, g7)


###
### MILTIPLICADORES DE IMPOSTOS (12)
###

## Coeficientes de renda ou requisitos de renda (12.1)

CTRIB = round((TRIB / VBP), digits = 4) %>% 
  as.vector()

matriz_TRIB = diag(CTRIB)

## Matriz de coeficientes de impostos indiretos (12.2)

TRIB1 = round((matriz_TRIB %*% B), digits = 4)

## Multiplicador de impostos indiretos Simples ou Gerador de impostos indiretos - MIIS (12.3)

MIIS = round((colSums(TRIB1)), digits = 4)

# Decomposicao intra e inter

intra1 = TRIB1[1:18, 1:18] %>% 
  colSums()

intra2 = TRIB1[19:36, 19:36] %>% 
  colSums()

intra = c(intra1, intra2)

## Multiplicador de impostos indiretos Tipo I - MIIT1 (12.4)

MIIT1 = round((MIIS / CTRIB), digits = 4)

## Multiplicador de impostos indiretos total truncado - MIITT (12.5)

trib_trunc = round(matriz_TRIB %*% B_fechado[1:n, 1:n], digits = 4)

MIITT = round(colSums(trib_trunc), digits = 4)

## Multiplicador de impostos indiretos Tipo II - MIIT2 (12.6)

MIIT2 = round(MIITT / CTRIB, digits = 4)

## Multiplicadores de impostos indiretos (12.7)

MultIMPI = data.frame("Setores" = SET$Setores,
                     "MIIS" = round(MIIS, digits = 4),
                     "Intra" = round(intra, digits = 4),
                     "Inter" = round(MIIS - intra, digits = 4),
                     "Direto" = round(CTRIB, digits = 4),
                     "indireto" = round(MIIS - CTRIB, digits = 4),
                     "MIITT" = round(MIITT, digits = 4),
                     "MIIT1" = round(MIIT1, digits = 4),
                     "MIIT2" = round(MIIT2, digits = 4))

rm(CTRIB, matriz_TRIB, TRIB1, MIIS, intra1, intra2, intra, MIIT1, trib_trunc, MIITT, MIIT2)

MultIMPI = MultIMPI %>% 
  mutate(MIIT1 = case_when(is.nan(MIIT1) ~ round(0, digits = 4),
                           TRUE ~ MIIT1),
         MIIT2 = case_when(MIIT2 == Inf ~ round(0, digits = 4),
                           TRUE ~ MIIT2))

## Tabela 8: Multiplicadores de impostos indiretos ()

t8 = MultIMPI[19:36, ] %>% 
  flextable() %>% 
  theme_zebra() %>% 
  align(align = "center", part = "header", i = 1) %>% 
  autofit(add_w = 0, add_h = 0) %>% 
  set_caption(caption = "Multiplicadores de renda simples e truncados") %>% 
  add_footer_lines("Fonte: Departamento de Contas Regionais e Finanças Públicas (DECRE-IMESC) - 2019. \nNota: MSR e MTRT por milhão de reais. A interpretação segue: por exemplo, uma variação de demanda de R$ 1.000.000 no setor Agro, gera 0,1977 de renda na economia.") %>% 
  flextable::footnote(i = 1, j = 2:5,
                      value = as_paragraph(c("O Multiplicador Simples de Renda (ou Gerador de Renda) apresenta o impacto total (direto + indireto) em termos de renda gerada na economia dada uma variação unitária da demanda final;", 
                                             "Multiplicador de Renda Tipo 1: efeito multiplicador na economia para cada unidade de renda gerada diretamente no setor específico;", 
                                             "O Multiplicador Total de Renda Truncado mostra os impactos direto, indireto e induzido, em termos de renda gerada na economia, dada uma variação unitária da demanda final;",
                                             "Multiplicador de Renda Tipo 2: efeito multiplicador na economia, considerando o efeito induzido, dada cada unidade de renda gerada no setor específico.")),
                      ref_symbols = c("a", "b", "c", "d"),
                      part = "header",
                      sep = "; ")

save_as_pptx("Multiplicadores de impostos indiretos - MA" = t8,
             path = "Anexo 6.Multiplicadores de impostos indiretos.pptx")

rm(t8)

## Grafico 8: Multiplicadores de impostos indiretos ()

g = MultIMPI %>% 
  head(n = 18) %>% 
  ggplot(aes(x = MIIT1, y = reorder(Setores, MIIT1, size = MIIT1))) +
  theme_minimal() + 
  geom_col(fill = "deepskyblue4") + 
  labs(title = "Multiplicador de impostos indiretos tipo 1 - modelo aberto",
       x = " ",
       y = " ") +
  theme(axis.text = element_text(color = "black", size = 15)) +
  geom_text(aes(label = round(MIIT1, digits = 4)), vjust = 0.5, hjust = 1, size = 5, color = "white") +
  scale_x_continuous(breaks = seq(0, 2.2, .2), limits = c(0, 2.2))

g1 = MultIMPI %>% 
  head(n = 18) %>% 
  ggplot(aes(x = MIIT2, y = reorder(Setores, MIIT2, size = MIIT2))) +
  theme_minimal() + 
  geom_col(fill = "deepskyblue4") + 
  labs(title = "Multiplicador de impostos indiretos tipo 2 - modelo fechado",
       x = " ",
       y = " ") +
  theme(axis.text = element_text(color = "black", size = 15)) +
  geom_text(aes(label = round(MIIT2, digits = 4)), vjust = 0.5, hjust = 1, size = 5, color = "white") +
  scale_x_continuous(breaks = seq(0, 8, 1), limits = c(0, 8))

#

g8 = g + g1

ggsave(filename = "Grafico 8.Multiplicadores de impostos indiretos.png",
       plot = g8,
       width = 18,
       height = 10)

rm(g, g1, g8)


###
### Índices de ligação (13)
###


## Índice de ligação para trás - Back Linkage (13.1)

bl = round((colMeans(B) / mean(B)), digits = 4)

## Índice de ligação para frente - Forward Linkages (13.2)

fl = round((rowMeans(B) / mean(B)), digits = 4)

## Índice de ligação para frente com Ghosh - (13.3)

SLG = rowSums(G)  # soma das linhas da matriz G
MLG = SLG / n  # média das somas da linhas da matriz G
Gstar = sum(G) / n ** 2  # média da matriz G
FLG = round((MLG / Gstar), digits = 4)  # índice de ligação p/ frente com a matriz G

## Índices de ligação (13.4)

Ind_Lig = cbind(SET, bl, fl, FLG) %>% 
  as.data.frame()

rm(bl, fl, SLG, MLG, Gstar, FLG)

colnames(Ind_Lig) = c("Setores", "Ligação para trás", "Ligação para frente", "Ligação p/ frente - Ghosh")

Ind_Lig = Ind_Lig %>% 
  mutate("Setor-Chave" = if_else(`Ligação para trás` > 1 & `Ligação p/ frente - Ghosh` > 1,
                               "Sim", "Não"))

# Tabela 8: Índices de ligação

t8 =  Ind_Lig[19:36, ] %>% 
  flextable() %>% 
  theme_zebra() %>% 
  align(align = "left", part = "header", i = 1) %>% 
  autofit(add_w = 0, add_h = 0) %>% 
  set_caption(caption = "Índices de Ligação e Setores-Chave") %>% 
  add_footer_lines("Departamento de Contas Regionais e Finanças Públicas (DECRE-IMESC) - 2019.") %>% 
  add_footer_lines("Nota: Setores-chaves baseados nos índices de ligação para trás e para frente-Ghosh.") %>% 
  flextable::footnote(i = 1, j = 2:4,
                      value = as_paragraph(c("Indica o quanto um setor demanda dos demais setores da economia. Índices maiores que 1 represetam forte ligaçã para trás, ou seja, o setor específico gera uma variação na demanda dos outros setores acima da média;",
                                             "Indica o quanto um setor é demandado pelos demais setores da economia. Índices maiores que 1 representam alta dependência do setor específico em relação ao restante da economia, ou seja, o setor particular tem uma interdepedência acima da média com outros setores;",
                                             "Esse índice tem a mesma interpretação do índice de ligação para frente.")),
                      ref_symbols = c("a", "b", "c"),
                      part = "header")

save_as_pptx("Índices de ligação - RBr" = t8,
             path = "Anexo 7.Índices de ligação.pptx")

rm(t8)


# Grafico 9: Índices de ligação

g9 = Ind_Lig %>% 
  head(n = 18) %>% 
ggplot(aes(x = `Ligação p/ frente - Ghosh`, y = `Ligação para trás`)) + 
  geom_point() + 
  theme_test() + 
  theme(plot.background = element_rect(fill = "gray87", colour = "gray87")) + 
  xlab(expression("Índice de ligação para frente"~"("*U[i]*")")) + 
  ylab(expression("Índice de ligação para trás"~"("*U[j]*")")) + 
  #ggtitle("Índices de Ligação e Setores-Chave") + 
  #labs(subtitle = "2019", 
       #caption = "Departamento de Contas Regionais e Finanças Públicas (DECRE-IMESC) - 2019. \nNota:Índice de ligação para frente calcuado com base na matriz inversa de Ghosh.") + 
  theme(plot.title = element_text(hjust = 0.5), 
        plot.subtitle = element_text(hjust = 0.5), 
        plot.caption = element_text(hjust = 0)) + 
  geom_text_repel(aes(label = Setores), size = 6)+ 
  scale_y_continuous(labels = scales::number_format(accuracy = 0.01), 
                     limits = c(0.4,1.6), 
                     breaks = seq(from = 0.2, to = 2, by = 0.2)) + 
  scale_x_continuous(labels = scales::number_format(accuracy = 0.01), 
                     limits = c(0.4,1.6), breaks = seq(from = 0.2, to = 2, by = 0.2)) + 
  geom_hline(yintercept=1, linetype="dashed", color = "black") + 
  geom_vline(xintercept=1, linetype="dashed", color = "black") + 
  annotate("text", x=1.59, y=1.6, label= "Setor-Chave", colour='black', size=7) + 
  annotate('text', x=0.49, y=1.6, label='Forte encadeamento para trás', colour='black', size=7) + 
  annotate('text', x=0.45, y=0.4, label='Fraco encadeamento', colour='black', size=7) + 
  annotate('text', x=1.5, y=0.4, label='Forte encadeamento para frente', colour='black', size=7) + 
  theme(axis.text = element_text(color = "black", size = 12))
  #ggimage::geom_image(aes(x = 1.2, 1.6),
   #                   image = "imesc4.png",
    #                  size = .3)
  
ggsave(filename = "Grafico 9.Índices de ligação.png",
       plot = g9,
       width = 19,
       height = 11)

rm(g9)


###
### Coeficientes de variação (14)
###


## coeficiente de variação da ligação para trás (14.1)

var1 = round((((1 / (n - 1)) * (rowSums((B - colMeans(B)) ** 2))) ** 0.5) / colMeans(B), digits = 4)

## coeficiente de variação da ligação para frente (14.2)

var2 = round((((1 / (n - 1)) * (colSums((B - rowMeans(B)) ** 2))) ** 0.5) / rowMeans(B), digits = 4)

## Coeficientes de variação (14.3)

CoefVar = cbind(SET, var1, var2) %>% 
  as.data.frame()

rm(var1, var2)

colnames(CoefVar) = c("Setores", "CV - trás",
                      "CV - frente")

# Tabela 9: Coeficientes de variação

t9 = CoefVar[19:36, ] %>% 
  flextable %>% 
  theme_zebra() %>% 
  align(align = "left", part = "header", i = 1) %>% 
  autofit(add_w = 0, add_h = 0) %>% 
  set_caption(caption = "Coeficientes de variação") %>% 
  add_footer_lines("Departamento de Contas Regionais e Finanças Públicas (DECRE-IMESC) - 2019.") %>% 
  flextable::footnote(i = 1, j = 2:3,
                      value = as_paragraph(c("Quanto menor for essa medida, maior será o número de setores atingidos pela variação na demanda final do setor específico;",
                                             "Quanto menor for essa medida, maior será o número de setores atendidos pelas vendas do setor específico.")),
                      ref_symbols = c("a", "b"),
                      part = "header")

save_as_pptx("Coeficientes de variação - RBr" = t9,
             path = "Anexo 8.Coeficientes de variação.pptx")

rm(t9)


###
### Campos de influencia (15)
###

ee = 0.001

E = matrix(0, ncol = n, nrow = n) 

SI = matrix(0, ncol = n, nrow = n)


for (i in 1:n) {
  for (j in 1:n) {
    E[i, j] = ee
    AE = A + E
    BE = solve(I - AE)
    FE = (BE - B) / ee
    FEq = FE * FE
    S = sum(FEq)
    SI[i, j] = S
    E[i, j] = 0
  }
}

# Grafico 10 - Campos de influencia (MA x MA)

sx = SET[1:18, ]

sy = SET[1:18, ]

data = expand.grid(x = sx, y = sy)

SI_1 = SI[1:18, 1:18]

# Grafico 11 - Campos de influencia (RBr x RBr)

sx = SET[19:36, ]

sy = SET[19:36, ]

data = expand.grid(x = sx, y = sy)

SI_2 = SI[19:36, 19:36]

#

g11 = ggplot(data,aes(x, y, fill = SI_1)) + 
  geom_tile() + 
  theme_bw() + 
  xlab(" ") + 
  ylab(" ") + 
  #ggtitle("Campo de Influência") + 
  #labs(subtitle = "2019", caption = "Departamento de Contas Regionais e Finanças Públicas (DECRE-IMESC) - 2019. \nNota: O campo de influência mostra como se distribuem as mudanças dos coeficientes diretos no sistema econômico como um todo, permitindo a determinação de quais relações entre os setores seriam mais importantes dentro do processo produtivo. Ou seja, \na determinação dos setores que apresentam um maior poder de influência sobre os demais, ou melhor, quais coeficientes que, alterados, teriam um maior impacto no sistema como um todo.") + 
  theme(plot.title = element_text(hjust = 0.5), 
        plot.subtitle = element_text(hjust = 0.5), 
        plot.caption = element_text(hjust = 0)) + 
  theme(axis.text.x = element_text(angle=45, vjust = 0.5), 
        axis.text.y = element_text(hjust = 0.5)) + 
  theme(legend.position = "none") + 
  scale_fill_distiller(palette = "Blues", trans = "reverse")

ggsave(filename = "Grafico 11.Campos de influencia (RBr).png",
       plot = g11,
       width = 18,
       height = 10,
       units = "in",
       dpi = 600)

rm(g10, g11, ee, E, AE, BE, FE, FEq, S, sx, sy, i, j, SI, SI_1, SI_2, data)


###
### Extração Hipotética (16)
###

## Extração da estrutura de compras (ligação para trás) – extração das colunas (16.1)

Extrac_tras = matrix(NA, ncol = 1, nrow = n)

## Extração da estrutura de vendas (ligação para frente) – extração das linhas (16.2)

Extrac_frente = matrix(NA, ncol = 1, nrow = n)

## Extração (16.3)

for (i in 1:n) {
  for (j in 1:n) {
    ABL = A
    ABL[, j] = 0
    BBL = solve(I - ABL)
    xbl = BBL %*% DF
    tbl = sum(VBP) - sum(xbl)
    Extrac_tras[j] = round(tbl)
    Extrac_tras_p = round((Extrac_tras / sum(VBP[19:36]) * 100), digits = 2)
    
    FFL = OF
    FFL[i, ] = 0
    GFL = solve(I - FFL)
    xfl = t(SP) %*% GFL
    tfl = sum(VBP) - sum(xfl)
    Extrac_frente[i] = round(tfl)
    Extrac_frente_p = round((Extrac_frente / sum(VBP[19:36]) * 100), digits = 2)
    
    Extracao = cbind(Extrac_tras, Extrac_tras_p, Extrac_frente, Extrac_frente_p)
    colnames(Extracao) = c("Extração p\ trás", "% do VBP1", "Extração p\ frente", "% do VBP2")
  }
}

Extracao = cbind(SET, Extracao) %>% 
  as.data.frame()

rm(ABL, BBL, xbl, tbl, Extrac_tras, Extrac_tras_p, FFL, GFL, xfl, tfl, Extrac_frente, Extrac_frente_p, i, j)

colnames(Extracao) = c("Setores", "Extração-Trás", "ET (%)", "Extração-Frente", "EF (%)")

# Tabela 10: Extração Hipotética - MA
# Anexo 8: Extração Hipotética - RBr

t11 = Extracao[19:36, ] %>% 
  flextable() %>% 
  colformat_double(j = c(2,4), big.mark = " ", digits = 0) %>% 
  colformat_double(j = c(3, 5), decimal.mark = ",", digits = 2, suffix = " %") %>% 
  theme_zebra() %>% 
  align(align = "left", part = "header", i = 1) %>% 
  autofit(add_w = 0, add_h = 0) %>% 
  set_caption(caption = "Extração Hipotética") %>% 
  add_footer_lines("Departamento de Contas Regionais e Finanças Públicas (DECRE-IMESC) - 2019.") %>% 
  flextable::footnote(i = 1, j = c(2, 4),
                      value = as_paragraph(c("Extração da estrutura de compras: é uma medida agregada de perda na economia - diminuição da produção total se o setor específico deixar de comprar dos outros setores, ou seja, representa a importância relativa do setor na interdependência econômica;",
                                             "Extração da estrutura de vendas: é uma medida agregada de perda na economia - diminuição da produção total se o setor específico deixar de vender para os outros setores.")),
                      ref_symbols = c("a", "b"),
                      part = "header")

save_as_pptx("Extração Hipotética - RBr" = t11,
             path = "Anexo 8.Extração Hipotética (RBr).pptx")

rm(t11, Extracao)



################################################################################
################################################################################
################################################################################
################################################################################