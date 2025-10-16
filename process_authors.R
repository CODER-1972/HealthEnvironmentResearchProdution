#!/usr/bin/env Rscript

options(repos = c(CRAN = "https://cloud.r-project.org"))
user_library <- Sys.getenv("R_LIBS_USER")
if (user_library != "" && !dir.exists(user_library)) {
  dir.create(user_library, recursive = TRUE, showWarnings = FALSE)
}
if (user_library != "") {
  .libPaths(unique(c(user_library, .libPaths())))
}

required_packages <- c("readxl", "writexl", "dplyr", "stringr", "purrr", "tidyr")

for (pkg in required_packages) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    message("A instalar pacote obrigatório: ", pkg)
    install.packages(pkg, lib = if (user_library == "") NULL else user_library, dependencies = TRUE)
  }
  suppressPackageStartupMessages(library(pkg, character.only = TRUE))
}

cat("=== Processamento de autores (Web of Science) ===\n")
folder_path <- readline(prompt = "Introduza o caminho completo da pasta que contém o ficheiro Excel: ")
folder_path <- str_trim(folder_path)
    
if (str_starts(folder_path, "\"") && str_ends(folder_path, "\"")) {
  folder_path <- str_sub(folder_path, 2, -2)
}
if (str_starts(folder_path, "'") && str_ends(folder_path, "'")) {
  folder_path <- str_sub(folder_path, 2, -2)
}
if (folder_path != "") {
  folder_path <- path.expand(folder_path)
  suppressWarnings({
    folder_path <- normalizePath(folder_path, winslash = "/", mustWork = FALSE)
  })
}

if (folder_path == "") {
  folder_path <- getwd()
  message("Nenhum caminho indicado. A pasta atual será utilizada: ", folder_path)
}

if (.Platform$OS.type != "windows" && str_detect(folder_path, "^[A-Za-z]:\\\\")) {
  stop(
    "O caminho indicado parece ser do Windows (", folder_path, "), mas o ambiente atual não consegue aceder a discos locais. ",
    "Carregue o ficheiro para a área de trabalho atual ou forneça um caminho válido neste sistema."
  )
}

if (!dir.exists(folder_path)) {
  stop(
    "A pasta indicada não existe: ", folder_path,
    ". Verifique se o caminho está correto e acessível a partir deste ambiente."
  )
}

excel_files <- list.files(folder_path, pattern = "\\.xlsx$|\\.xls$", ignore.case = TRUE, full.names = TRUE)
if (length(excel_files) == 0) {
  stop("Não foram encontrados ficheiros Excel na pasta indicada.")
}

if (length(excel_files) == 1) {
  excel_path <- excel_files[[1]]
  message("Ficheiro encontrado: ", basename(excel_path))
} else {
  cat("Foram encontrados os seguintes ficheiros Excel:\n")
  for (i in seq_along(excel_files)) {
    cat(sprintf("[%d] %s\n", i, basename(excel_files[[i]])))
  }
  selection <- readline(prompt = "Indique o número do ficheiro a utilizar: ")
  selection <- suppressWarnings(as.integer(selection))
  if (is.na(selection) || selection < 1 || selection > length(excel_files)) {
    stop("Seleção inválida.")
  }
  excel_path <- excel_files[[selection]]
}

message("A ler o ficheiro: ", excel_path)
data <- read_excel(excel_path)

strip_diacritics <- function(values) {
  if (length(values) == 0) {
    return(values)
  }
  result <- suppressWarnings(iconv(values, from = "UTF-8", to = "ASCII//TRANSLIT"))
  ifelse(is.na(result), values, result)
}

normalize_name <- function(x) {
  if (is.null(x)) {
    return("")
  }
  x <- strip_diacritics(tolower(x))
  str_replace_all(x, "[^a-z0-9]", "")
}

find_column <- function(df, candidates) {
  normalized_cols <- vapply(names(df), normalize_name, character(1))
  for (candidate in candidates) {
    candidate_norm <- normalize_name(candidate)
    matches <- names(df)[normalized_cols == candidate_norm]
    if (length(matches) > 0) {
      return(matches[[1]])
    }
  }
  NULL
}

split_authors <- function(value) {
  if (length(value) == 0 || all(is.na(value))) {
    return(character())
  }
  value <- str_replace_all(value, "\\r\\n|\\n", "; ")
  tokens <- str_split(value, ";\\s*")[[1]]
  tokens <- str_trim(tokens)
  tokens[tokens != ""]
}

make_author_key <- function(name) {
  if (length(name) == 0) {
    return(character())
  }
  vapply(
    name,
    function(single) {
      if (is.null(single) || length(single) == 0) {
        return(NA_character_)
      }
      single <- as.character(single)[1]
      if (is.na(single)) {
        return(NA_character_)
      }
      cleaned <- strip_diacritics(single)
      cleaned <- str_replace_all(cleaned, "\\r\\n|\\n", " ")
      cleaned <- str_replace_all(cleaned, "\\.+", "")
      cleaned <- str_squish(cleaned)
      if (cleaned == "") {
        return(NA_character_)
      }
      cleaned_upper <- str_to_upper(cleaned)
      if (str_detect(cleaned_upper, ",")) {
        parts <- str_split_fixed(cleaned_upper, ",", 2)
        surname <- str_squish(parts[, 1])
        given <- str_squish(parts[, 2])
      } else {
        pieces <- str_split(cleaned_upper, "\\s+")[[1]]
        pieces <- pieces[pieces != ""]
        if (length(pieces) == 0) {
          return(NA_character_)
        }
        surname <- pieces[1]
        if (length(pieces) > 1) {
          given <- str_trim(paste(pieces[-1], collapse = " "))
        } else {
          given <- ""
        }
      }
      if (surname == "") {
        return(NA_character_)
      }
      given_parts <- str_split(given, "\\s+")[[1]]
      given_parts <- given_parts[given_parts != ""]
      initials <- if (length(given_parts) == 0) "" else paste0(str_sub(given_parts, 1, 1), collapse = "")
      key <- str_trim(paste(surname, initials))
      if (key == "") NA_character_ else key
    },
    character(1),
    USE.NAMES = FALSE
  )
}

make_author_key_from_token <- function(token) {
  if (is.null(token) || length(token) == 0) {
    return(NA_character_)
  }
  token <- token[1]
  if (is.na(token) || token == "") {
    return(NA_character_)
  }
  token <- strip_diacritics(token)
  token <- str_replace_all(token, "\\r\\n|\\n", " ")
  token <- str_replace_all(token, "\\.+", "")
  token <- str_squish(token)
  if (token == "") {
    return(NA_character_)
  }
  upper <- str_to_upper(token)
  if (!str_detect(upper, ",")) {
    pieces <- str_split(upper, "\\s+", n = 2)[[1]]
    if (length(pieces) >= 2) {
      upper <- paste(pieces[1], pieces[2], sep = ", ")
    } else {
      upper <- paste0(pieces[1], ",")
    }
  }
  key <- make_author_key(upper)
  if (length(key) == 0) NA_character_ else key[[1]]
}

parse_orcid_entries <- function(value) {
  if (is.null(value) || all(is.na(value))) {
    return(tibble(AutorKey = character(), ORCID = character()))
  }
  value <- str_replace_all(value, "\\r\\n|\\n", " ")
  tokens <- str_split(value, ";\\s*")[[1]]
  tokens <- str_trim(tokens)
  tokens <- tokens[tokens != ""]
  if (length(tokens) == 0) {
    return(tibble(AutorKey = character(), ORCID = character()))
  }
  orcid_pattern <- "\\\\d{4}-\\\\d{4}-\\\\d{4}-[\\\\dX]{4}"
  map_dfr(tokens, function(token) {
    orcid <- str_extract(token, orcid_pattern)
    if (is.na(orcid)) {
      return(tibble(AutorKey = character(), ORCID = character()))
    }
    name_part <- NA_character_
    if (str_detect(token, "/")) {
      name_part <- str_trim(str_split(token, "/")[[1]][1])
    } else {
      without_orcid <- str_trim(str_remove(token, orcid))
      without_orcid <- str_remove(without_orcid, "[,/]+$")
      if (without_orcid != "") {
        name_part <- without_orcid
      }
    }
    autor_key <- if (!is.na(name_part) && name_part != "") make_author_key(name_part) else NA_character_
    tibble(AutorKey = autor_key, ORCID = orcid)
  }) %>%
    distinct()
}

parse_affiliation_entries <- function(value, row_id) {
  if (is.null(value) || is.na(value) || str_trim(value) == "") {
    return(tibble(RowID = integer(), AutorKey = character(), Affiliation = character()))
  }
  cleaned <- str_replace_all(value, "\\r\\n|\\n", " ")
  pattern <- "\\\\[(.*?)\\\\]\\\\s*([^\\\\[]+)"
  matches <- str_match_all(cleaned, pattern)[[1]]
  if (nrow(matches) == 0) {
    affiliation <- str_squish(cleaned)
    if (affiliation == "") {
      return(tibble(RowID = integer(), AutorKey = character(), Affiliation = character()))
    }
    return(tibble(RowID = row_id, AutorKey = NA_character_, Affiliation = affiliation))
  }
  map_dfr(seq_len(nrow(matches)), function(i) {
    author_segment <- matches[i, 2]
    affiliation <- str_squish(matches[i, 3])
    affiliation <- str_remove(affiliation, ";\\s*$")
    if (affiliation == "") {
      return(tibble(RowID = integer(), AutorKey = character(), Affiliation = character()))
    }
    authors <- str_split(author_segment, ";\\s*")[[1]]
    authors <- str_trim(authors)
    authors <- authors[authors != ""]
    if (length(authors) == 0) {
      return(tibble(RowID = row_id, AutorKey = NA_character_, Affiliation = affiliation))
    }
    keys <- map_chr(authors, make_author_key_from_token)
    if (all(is.na(keys))) {
      tibble(RowID = row_id, AutorKey = NA_character_, Affiliation = affiliation)
    } else {
      tibble(RowID = row_id, AutorKey = keys[!is.na(keys)], Affiliation = affiliation)
    }
  }) %>%
    distinct()
}

combine_orcid <- function(values) {
  values <- values %>%
    discard(~ is.na(.x) || .x == "") %>%
    unique()
  if (length(values) == 0) {
    return(NA_character_)
  }
  if (length(values) > 1) {
    warning("Foram encontrados múltiplos ORCID para um mesmo autor. Será utilizado o primeiro valor disponível.")
  }
  values[[1]]
}

split_affiliations <- function(values) {
  if (length(values) == 0) {
    return(NA_character_)
  }
  values <- values %>%
    discard(~ is.na(.x) || .x == "")
  if (length(values) == 0) {
    return(NA_character_)
  }
  tokens <- values %>%
    str_split(pattern = "[;|]\\\\s*") %>%
    flatten_chr() %>%
    str_trim() %>%
    discard(~ .x == "") %>%
    unique()
  if (length(tokens) == 0) {
    return(NA_character_)
  }
  paste(tokens, collapse = "; ")
}

full_name_col <- find_column(
  data,
  c(
    "author full names", "autores nomes completos", "af", "authors", "autores", "author", "au", "nome", "name"
  )
)
if (is.null(full_name_col)) {
  stop(
    "Não foi possível identificar a coluna com os nomes dos autores (por exemplo, 'Author Full Names' ou 'AF')."
  )
}

author_raw <- data[[full_name_col]]
author_rows <- tibble(
  RowID = seq_along(author_raw),
  Autores = map(author_raw, split_authors)
) %>%
  mutate(Autores = map(Autores, ~ .x[.x != ""])) %>%
  unnest(cols = Autores) %>%
  rename(Autor = Autores) %>%
  mutate(
    Autor = str_squish(Autor),
    AutorKey = make_author_key(Autor)
  ) %>%
  filter(!is.na(Autor) & Autor != "")

if (nrow(author_rows) == 0) {
  stop("Não foram encontrados autores no ficheiro fornecido.")
}

author_counts <- author_rows %>%
  count(RowID, name = "AutorCount")

author_rows <- author_rows %>%
  left_join(author_counts, by = "RowID")

orcid_col <- find_column(data, c("author identifiers", "identificadores de autores", "orcid", "orcid id", "orcidid", "oi"))
affiliation_col <- find_column(
  data,
  c(
    "addresses", "address", "affiliations", "affiliation", "c1", "filiacao", "filiação", "instituicao", "institution"
  )
)

if (is.null(orcid_col)) {
  message("Aviso: coluna de ORCID não encontrada. Será criado um campo vazio.")
}

orcid_map <- if (!is.null(orcid_col)) {
  tibble(RowID = seq_len(nrow(data)), Valor = data[[orcid_col]]) %>%
    mutate(Registos = map(Valor, parse_orcid_entries)) %>%
    select(RowID, Registos) %>%
    unnest(Registos)
} else {
  tibble(RowID = integer(), AutorKey = character(), ORCID = character())
}

if (!is.null(orcid_col) && nrow(orcid_map) == 0) {
  message("Aviso: coluna de ORCID encontrada mas não foram identificados identificadores válidos.")
}

orcid_specific <- orcid_map %>%
  filter(!is.na(AutorKey) & !is.na(ORCID)) %>%
  distinct(RowID, AutorKey, ORCID)

orcid_unspecified <- orcid_map %>%
  filter((is.na(AutorKey) | AutorKey == "") & !is.na(ORCID)) %>%
  group_by(RowID) %>%
  summarise(ORCIDLista = list(unique(ORCID)), .groups = "drop")

authors_with_orcid <- author_rows %>%
  left_join(orcid_specific, by = c("RowID", "AutorKey"))

if (nrow(orcid_unspecified) > 0) {
  authors_with_orcid <- authors_with_orcid %>%
    left_join(orcid_unspecified, by = "RowID") %>%
    mutate(
      ORCID = if_else(
        is.na(ORCID) & map_int(ORCIDLista, ~ if (is.null(.x) || all(is.na(.x))) 0L else length(.x)) == 1 & AutorCount == 1,
        map_chr(
          ORCIDLista,
          ~ if (is.null(.x) || length(.x) == 0 || all(is.na(.x))) NA_character_ else .x[[1]]
        ),
        ORCID
      )
    ) %>%
    select(-ORCIDLista)
}

if (is.null(affiliation_col)) {
  message("Aviso: coluna de filiação/instituições não encontrada. Será criado um campo vazio.")
  affiliation_map <- tibble(RowID = integer(), AutorKey = character(), Affiliation = character())
} else {
  affiliation_map <- tibble(RowID = seq_len(nrow(data)), Valor = data[[affiliation_col]]) %>%
    mutate(Registos = map2(Valor, RowID, parse_affiliation_entries)) %>%
    select(Registos) %>%
    unnest(Registos)
  if (nrow(affiliation_map) == 0) {
    message("Aviso: coluna de filiação encontrada mas não foi possível associar instituições específicas aos autores.")
  }
}

affiliation_specific <- affiliation_map %>%
  filter(!is.na(AutorKey) & AutorKey != "")

affiliation_unspecified <- affiliation_map %>%
  filter(is.na(AutorKey) | AutorKey == "")

if (nrow(affiliation_unspecified) > 0) {
  expanded_unspecified <- affiliation_unspecified %>%
    distinct(RowID, Affiliation) %>%
    inner_join(authors_with_orcid %>% distinct(RowID, AutorKey), by = "RowID")
  affiliation_combined <- bind_rows(affiliation_specific, expanded_unspecified)
} else {
  affiliation_combined <- affiliation_specific
}

authors_combined <- authors_with_orcid %>%
  left_join(affiliation_combined, by = c("RowID", "AutorKey"))

result <- authors_combined %>%
  group_by(Autor) %>%
  summarise(
    ORCID = combine_orcid(ORCID),
    Instituicoes = split_affiliations(Affiliation),
    .groups = "drop"
  ) %>%
  mutate(
    ORCID = replace_na(ORCID, ""),
    Instituicoes = replace_na(Instituicoes, "")
  ) %>%
  arrange(Autor)

output_path <- file.path(folder_path, "autores_unicos.xlsx")
write_xlsx(result, output_path)

message("Ficheiro criado: ", output_path)
