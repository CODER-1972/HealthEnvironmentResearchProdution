#!/usr/bin/env Rscript

cat("=== Descarregar process_authors.R ===\n")

# Função para tentar derivar o URL bruto a partir do remote do GitHub
derive_default_url <- function() {
  remote <- tryCatch(
    system("git config --get remote.origin.url", intern = TRUE),
    error = function(e) character()
  )
  if (length(remote) == 0) {
    return("")
  }
  remote <- remote[[1]]
  remote <- trimws(remote)
  if (remote == "") {
    return("")
  }
  if (grepl("^git@github.com:", remote)) {
    remote <- sub("^git@github.com:", "https://github.com/", remote)
  }
  remote <- sub("\\.git$", "", remote)
  if (!grepl("^https?://github.com/", remote, ignore.case = TRUE)) {
    return("")
  }
  remote_path <- sub("^https?://github.com/", "", remote, ignore.case = TRUE)
  pieces <- strsplit(remote_path, "/", fixed = TRUE)[[1]]
  if (length(pieces) < 2) {
    return("")
  }
  owner <- pieces[1]
  repo <- pieces[2]
  branch <- Sys.getenv("PROCESS_AUTHORS_BRANCH")
  if (branch == "") {
    branch <- tryCatch(
      system("git rev-parse --abbrev-ref HEAD", intern = TRUE),
      error = function(e) character()
    )
    if (length(branch) == 0 || branch[[1]] %in% c("HEAD", "")) {
      branch <- "main"
    } else {
      branch <- branch[[1]]
    }
  }
  paste0("https://raw.githubusercontent.com/", owner, "/", repo, "/", branch, "/process_authors.R")
}

default_url <- Sys.getenv("PROCESS_AUTHORS_DOWNLOAD_URL")
if (default_url == "") {
  default_url <- derive_default_url()
}

if (default_url != "") {
  message("URL predefinido identificado: ", default_url)
  message("Pode alterar o URL durante o próximo passo, se necessário.")
} else {
  message("Não foi possível determinar automaticamente o URL do script.")
}

input_prompt <- if (default_url == "") {
  "Introduza o URL completo para descarregar o script process_authors.R: "
} else {
  paste0(
    "Introduza o URL completo para descarregar o script process_authors.R\n",
    "(pressione Enter para utilizar o valor predefinido): "
  )
}

download_url <- readline(prompt = input_prompt)
download_url <- trimws(download_url)
if (download_url == "") {
  if (default_url == "") {
    stop("URL não fornecido. Não é possível continuar sem um endereço de origem.")
  }
  download_url <- default_url
}

message("URL selecionado: ", download_url)

default_dest <- file.path(getwd(), "process_authors.R")
dest_prompt <- paste0(
  "Indique o caminho de destino do ficheiro (Enter para utilizar ",
  default_dest,"): "
)

destination <- readline(prompt = dest_prompt)
destination <- trimws(destination)
if (destination == "") {
  destination <- default_dest
}

destination_dir <- dirname(destination)
if (!dir.exists(destination_dir)) {
  message("A criar diretório: ", destination_dir)
  dir.create(destination_dir, recursive = TRUE, showWarnings = FALSE)
}

message("A descarregar o script...")

success <- tryCatch(
  {
    utils::download.file(download_url, destfile = destination, mode = "wb", quiet = FALSE)
    TRUE
  },
  warning = function(w) {
    message("Aviso durante a transferência: ", conditionMessage(w))
    FALSE
  },
  error = function(e) {
    message("Erro durante a transferência: ", conditionMessage(e))
    FALSE
  }
)

if (!success) {
  stop("Não foi possível descarregar o script. Verifique o URL e a ligação à internet.")
}

message("Script guardado em: ", destination)
