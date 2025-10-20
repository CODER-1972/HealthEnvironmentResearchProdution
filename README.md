# Processamento de Autores de ficheiros Web of Science

Este repositório contém scripts em R para consolidar autores, ORCID e instituições
a partir de ficheiros Excel exportados do Web of Science.

## Scripts disponíveis

- `process_authors.R` – Lê um ficheiro Excel do Web of Science, agrega nomes
  únicos de autores, os respetivos ORCID (quando disponíveis; múltiplos
  identificadores para o mesmo autor são apresentados separados por `;`) e as
  instituições associadas. ORCID que existam no ficheiro mas que não possuam
  nome associado são incluídos no resultado com a coluna de autor vazia, para
  que nenhum identificador se perca. Além da folha principal com autores,
  passa a gerar uma folha adicional que agrupa entradas que partilham o mesmo
  ORCID, apresentando os nomes ordenados pela frequência com que surgem no
  ficheiro de origem.
- `download_process_authors.R` – Auxilia a transferência do script principal
  (`process_authors.R`) para o seu ambiente de trabalho, deduzindo
  automaticamente o URL bruto do GitHub sempre que possível.

## Como descarregar o script

Se estiver a trabalhar num ambiente como o Posit Cloud (ou qualquer outra
máquina onde ainda não tenha o ficheiro), pode executar:

```r
source("download_process_authors.R")
```

O utilitário irá:

1. Tentar identificar automaticamente o URL bruto do repositório GitHub.
2. Perguntar-lhe qual o URL a utilizar (permitindo aceitar o valor sugerido).
3. Solicitar o caminho de destino do ficheiro `process_authors.R`.
4. Transferir o script para o local indicado, criando diretórios em falta.

Também pode definir manualmente as seguintes variáveis de ambiente antes de
executar o utilitário:

- `PROCESS_AUTHORS_DOWNLOAD_URL` – URL completo para o ficheiro `process_authors.R`.
- `PROCESS_AUTHORS_BRANCH` – Nome do *branch* a utilizar ao deduzir o URL a
  partir do `git remote` (por defeito tenta usar o branch atual e, em seguida,
  `main`).

## Execução do processamento

```bash
Rscript process_authors.R
```

Será solicitada a pasta onde se encontra o ficheiro Excel exportado do Web of
Science, e o script gera `autores_unicos.xlsx` com duas folhas:

1. **Autores** – Lista consolidada de autores, ORCID e instituições.
2. **ORCID Agrupados** – Entradas com o mesmo ORCID são agregadas, listando os
   nomes correspondentes (ordenados do mais frequente para o menos frequente),
   repetindo o identificador ORCID pela mesma ordem e alinhando as instituições
   associadas a cada autor.

> ℹ️ **Importante:** Em ambientes remotos (por exemplo, Posit Cloud) não é
> possível aceder diretamente a discos locais como `C:\\Users\\...`. Carregue o
> ficheiro para o projeto remoto e indique o caminho correspondente (por
> exemplo, `Dados/`). Caminhos de Windows só devem ser utilizados quando o
> script é executado numa máquina Windows com acesso a esse diretório.
