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
  value <- str_replace_all(value, "\\r\\n|\\n", "; ")
  tokens <- str_split(value, ";\\s*")[[1]]
  tokens <- str_trim(tokens)
  tokens <- tokens[tokens != ""]
  if (length(tokens) == 0) {
    return(tibble(AutorKey = character(), ORCID = character()))
  }
  orcid_pattern <- regex(
    "(?:https?://orcid\\.org/)?([0-9]{4}[-\\s]?[0-9]{4}[-\\s]?[0-9]{4}[-\\s]?[0-9X]{4})",
    ignore_case = TRUE
  )
  map_dfr(tokens, function(token) {
    matches <- str_match_all(token, orcid_pattern)[[1]]
    if (nrow(matches) == 0) {
      return(tibble(AutorKey = character(), ORCID = character()))
    }
    orcids <- matches[, ncol(matches)]
    orcids <- str_replace_all(orcids, "[\\s−–—]", "-")
    orcids <- str_replace_all(orcids, "-+", "-")
    orcids <- str_to_upper(orcids)
    name_part <- NA_character_
    if (str_detect(token, "/")) {
      name_part <- str_split(token, "\\s*/\\s*", n = 2)[[1]]
      name_part <- str_trim(name_part[1])
    }
    if (is.na(name_part) || name_part == "") {
      without_orcid <- token
      for (current in matches[, 1]) {
        without_orcid <- str_replace(without_orcid, fixed(current), " ")
      }
      without_orcid <- str_remove_all(without_orcid, "https?://orcid\\.org/?")
      without_orcid <- str_remove(without_orcid, "(?i)orcid[:]?")
      without_orcid <- str_squish(without_orcid)
      without_orcid <- str_remove(without_orcid, "[,/;]+$")
      if (without_orcid != "") {
        name_part <- without_orcid
      }
    }
    autor_key <- if (!is.na(name_part) && name_part != "") make_author_key(name_part) else NA_character_
    if (length(autor_key) == 0) {
      autor_key <- NA_character_
    }
    tibble(AutorKey = rep_len(autor_key, length(orcids)), ORCID = orcids)
  }) %>%
    distinct()
}

parse_affiliation_entries <- function(value, row_id) {
  if (is.null(value) || is.na(value) || str_trim(value) == "") {
    return(tibble(RowID = integer(), AutorKey = character(), Affiliation = character()))
  }
  cleaned <- str_replace_all(value, "\\r\\n|\\n", " ")
  pattern <- "\\[(.*?)\\]\\s*([^\\[]+)"
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
  paste(values, collapse = "; ")
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

extract_surname_initial <- function(name) {
  if (is.null(name) || length(name) == 0) {
    return(NA_character_)
  }
  single <- as.character(name)[1]
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
  upper <- str_to_upper(cleaned)
  if (str_detect(upper, ",")) {
    parts <- str_split_fixed(upper, ",", 2)
    surname <- str_squish(parts[, 1])
    given <- str_squish(parts[, 2])
  } else {
    pieces <- str_split(upper, "\\s+")[[1]]
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
  initial <- if (length(given_parts) == 0) "" else str_sub(given_parts[1], 1, 1)
  key <- str_trim(paste(surname, initial))
  if (key == "") NA_character_ else key
}

collect_affiliation_tokens <- function(values) {
  if (length(values) == 0) {
    return(character())
  }
  values <- values[!is.na(values)]
  if (length(values) == 0) {
    return(character())
  }
  tokens <- values %>%
    str_replace_all("\\r\\n|\\n", "; ") %>%
    str_split(pattern = "[;|]\\s*") %>%
    flatten_chr() %>%
    str_trim()
  tokens[tokens != ""]
}

normalize_institution_token <- function(value) {
  if (is.null(value) || is.na(value)) {
    return("")
  }
  normalized <- strip_diacritics(value)
  normalized <- str_to_lower(normalized)
  normalized <- str_replace_all(normalized, "[^a-z0-9]+", " ")
  str_squish(normalized)
}

group_blocks_by_edges <- function(total_blocks, edges_df) {
  if (total_blocks == 0) {
    return(list())
  }

  adjacency <- vector("list", total_blocks)
  if (nrow(edges_df) > 0) {
    for (i in seq_len(nrow(edges_df))) {
      from <- edges_df$From[[i]]
      to <- edges_df$To[[i]]
      adjacency[[from]] <- unique(c(adjacency[[from]], to))
      adjacency[[to]] <- unique(c(adjacency[[to]], from))
    }
  }

  visited <- rep(FALSE, total_blocks)
  components <- list()

  for (i in seq_len(total_blocks)) {
    if (visited[i]) {
      next
    }
    stack <- i
    component <- integer()

    while (length(stack) > 0) {
      current <- stack[[length(stack)]]
      stack <- stack[-length(stack)]
      if (visited[current]) {
        next
      }
      visited[current] <- TRUE
      component <- unique(c(component, current))
      neighbours <- adjacency[[current]]
      if (length(neighbours) > 0) {
        stack <- unique(c(stack, setdiff(neighbours, component)))
      }
    }

    components[[length(components) + 1]] <- sort(component)
  }

  components
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

orcid_col <- find_column(
  data,
  c(
    "author identifiers",
    "identificadores de autores",
    "orcid",
    "orcid id",
    "orcidid",
    "orcids",
    "oi",
    "researcher ids / orcid (wos)",
    "researcher ids / orcid",
    "researcher ids",
    "researcherid",
    "researcherid numbers"
  )
)
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

assigned_orcids <- authors_with_orcid %>%
  filter(!is.na(ORCID) & ORCID != "") %>%
  distinct(RowID, ORCID)

orcid_orphans <- orcid_map %>%
  filter((is.na(AutorKey) | AutorKey == "") & !is.na(ORCID) & ORCID != "") %>%
  distinct(RowID, ORCID) %>%
  anti_join(assigned_orcids, by = c("RowID", "ORCID")) %>%
  distinct(ORCID)

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

author_summary <- authors_combined %>%
  group_by(Autor) %>%
  summarise(
    Ordem = suppressWarnings(min(RowID, na.rm = TRUE)),
    ORCID = combine_orcid(ORCID),
    Instituicoes = split_affiliations(Affiliation),
    .groups = "drop"
  ) %>%
  mutate(
    ORCID = replace_na(ORCID, ""),
    Instituicoes = replace_na(Instituicoes, ""),
    Ordem = ifelse(is.infinite(Ordem), NA_real_, Ordem)
  ) %>%
  arrange(Ordem, Autor)

author_summary <- author_summary %>%
  mutate(OrderIndex = row_number())
author_summary <- author_summary %>%
  mutate(Ordem = if_else(is.na(Ordem), as.numeric(OrderIndex), Ordem))

if (nrow(orcid_orphans) > 0) {
  max_ordem <- if (nrow(author_summary) == 0) 0 else max(author_summary$Ordem, na.rm = TRUE)
  max_index <- if (nrow(author_summary) == 0) 0 else max(author_summary$OrderIndex, na.rm = TRUE)
  orphan_rows <- orcid_orphans %>%
    mutate(
      Autor = "",
      Instituicoes = "",
      Ordem = max_ordem + seq_len(n()),
      OrderIndex = max_index + seq_len(n())
    ) %>%
    select(Autor, ORCID, Instituicoes, Ordem, OrderIndex)
  author_summary <- bind_rows(author_summary, orphan_rows)
}

author_summary <- author_summary %>%
  arrange(OrderIndex)

result <- author_summary %>%
  select(Autor, ORCID, Instituicoes) %>%
  mutate(
    AutorOrdenacao = str_trim(coalesce(Autor, ""))
  ) %>%
  arrange(AutorOrdenacao == "", AutorOrdenacao) %>%
  select(-AutorOrdenacao)

author_rows_for_groups <- author_summary %>%
  filter(Autor != "") %>%
  mutate(
    ORCIDOriginal = ORCID,
    ORCIDTokens = map(
      ORCID,
      ~ {
        if (is.null(.x) || .x == "") {
          character()
        } else {
          tokens <- str_split(.x, "\\s*;\\s*")[[1]]
          tokens <- str_trim(tokens)
          tokens[tokens != ""]
        }
      }
    )
  )

without_orcid <- author_rows_for_groups %>%
  filter(map_int(ORCIDTokens, length) == 0) %>%
  mutate(
    Instituicoes = replace_na(Instituicoes, ""),
    Detalhes = map2(
      Autor,
      Instituicoes,
      ~ tibble(
        Autor = .x,
        ORCIDTokens = list(character()),
        Instituicoes = .y
      )
    )
  ) %>%
  transmute(
    Ordem = OrderIndex,
    Autores = Autor,
    ORCIDs = "",
    Instituicoes = Instituicoes,
    Detalhes = Detalhes
  )

with_orcid <- author_rows_for_groups %>%
  filter(map_int(ORCIDTokens, length) > 0) %>%
  mutate(
    ORCIDTokens = map(ORCIDTokens, ~ unique(.x[!is.na(.x) & .x != ""])),
    AutorID = row_number(),
    Instituicoes = replace_na(Instituicoes, "")
  )

group_components_by_orcid <- function(df) {
  if (nrow(df) == 0) {
    return(list())
  }

  token_links <- df %>%
    select(AutorID, ORCIDTokens) %>%
    unnest_longer(ORCIDTokens, values_to = "Token") %>%
    distinct(AutorID, Token)

  if (nrow(token_links) == 0) {
    return(as.list(seq_len(nrow(df))))
  }

  token_lookup <- split(token_links$AutorID, token_links$Token)
  visited <- rep(FALSE, nrow(df))
  components <- list()

  for (i in seq_len(nrow(df))) {
    if (visited[i]) {
      next
    }
    stack <- i
    comp_ids <- integer()

    while (length(stack) > 0) {
      current <- stack[[length(stack)]]
      stack <- stack[-length(stack)]
      if (visited[current]) {
        next
      }
      visited[current] <- TRUE
      comp_ids <- unique(c(comp_ids, current))
      tokens_current <- df$ORCIDTokens[[current]]
      if (length(tokens_current) == 0) {
        next
      }
      for (token in tokens_current) {
        linked <- token_lookup[[token]]
        if (length(linked) == 0) {
          next
        }
        stack <- unique(c(stack, setdiff(linked, comp_ids)))
      }
    }

    components[[length(components) + 1]] <- comp_ids
  }

  components
}

component_indices <- group_components_by_orcid(with_orcid)

grouped_orcid <- if (length(component_indices) == 0) {
  tibble(
    Ordem = numeric(),
    Autores = character(),
    ORCIDs = character(),
    Instituicoes = character(),
    Detalhes = list()
  )
} else {
  map_dfr(component_indices, function(idx) {
    comp_authors <- with_orcid %>%
      filter(AutorID %in% idx) %>%
      arrange(OrderIndex)

    nomes <- comp_authors$Autor
    orcid_por_autor <- map_chr(
      comp_authors$ORCIDTokens,
      ~ if (length(.x) == 0) "" else paste(.x, collapse = "; ")
    )
    instituicoes <- replace_na(comp_authors$Instituicoes, "")

    tibble(
      Ordem = suppressWarnings(min(comp_authors$OrderIndex, na.rm = TRUE)),
      Autores = str_squish(paste(nomes, collapse = "; ")),
      ORCIDs = str_squish(paste(orcid_por_autor, collapse = "; ")),
      Instituicoes = str_squish(paste(instituicoes, collapse = "; ")),
      Detalhes = list(
        tibble(
          Autor = nomes,
          ORCIDTokens = comp_authors$ORCIDTokens,
          Instituicoes = instituicoes
        )
      )
    )
  })
}

orcid_groups_enriched <- bind_rows(grouped_orcid, without_orcid) %>%
  mutate(
    PrimeiroAutor = str_trim(coalesce(str_split_fixed(Autores, ";", 2)[, 1], ""))
  ) %>%
  arrange(PrimeiroAutor == "", PrimeiroAutor, Autores) %>%
  select(-PrimeiroAutor)

orcid_groups <- orcid_groups_enriched %>%
  select(Autores, ORCIDs, Instituicoes)

blocks <- orcid_groups_enriched %>%
  mutate(
    BlockID = row_number(),
    Detalhes = map(Detalhes, function(tbl) {
      if (is.null(tbl) || nrow(tbl) == 0) {
        tibble(
          Autor = character(),
          ORCIDTokens = list(),
          Instituicoes = character(),
          AutorIndex = integer()
        )
      } else {
        tbl %>%
          mutate(
            ORCIDTokens = map(ORCIDTokens, function(tokens) {
              if (is.null(tokens) || length(tokens) == 0) {
                character()
              } else {
                values <- as.character(tokens)
                values <- values[!is.na(values) & values != ""]
                unique(values)
              }
            }),
            Instituicoes = replace_na(Instituicoes, ""),
            AutorIndex = seq_len(n())
          )
      }
    })
  )

block_author_details <- blocks %>%
  select(BlockID, Detalhes) %>%
  unnest(Detalhes)

if (nrow(block_author_details) == 0) {
  block_author_details <- tibble(
    BlockID = integer(),
    Autor = character(),
    ORCIDTokens = list(),
    Instituicoes = character(),
    AutorIndex = integer()
  )
}

block_author_details <- block_author_details %>%
  mutate(
    SurnameInitial = map_chr(Autor, extract_surname_initial),
    InstituicaoTokens = map(Instituicoes, collect_affiliation_tokens)
  )

author_similarity_tokens <- block_author_details %>%
  filter(!is.na(SurnameInitial) & SurnameInitial != "") %>%
  mutate(InstituicaoTokens = map(InstituicaoTokens, ~ unique(.x))) %>%
  filter(map_int(InstituicaoTokens, length) > 0) %>%
  unnest(InstituicaoTokens, values_to = "InstituicaoToken") %>%
  mutate(
    InstituicaoKey = normalize_institution_token(InstituicaoToken)
  ) %>%
  filter(InstituicaoKey != "") %>%
  distinct(BlockID, SurnameInitial, InstituicaoKey)

similarity_edges <- if (nrow(author_similarity_tokens) == 0) {
  tibble(From = integer(), To = integer())
} else {
  author_similarity_tokens %>%
    inner_join(author_similarity_tokens, by = c("SurnameInitial", "InstituicaoKey"), suffix = c("_1", "_2")) %>%
    filter(BlockID_1 < BlockID_2) %>%
    transmute(From = BlockID_1, To = BlockID_2) %>%
    distinct()
}

components_blocks <- group_blocks_by_edges(nrow(blocks), similarity_edges)

surname_initial_groups <- map_dfr(components_blocks, function(component_ids) {
  if (length(component_ids) == 0) {
    return(tibble(Ordem = numeric(), Autores = character(), ORCIDs = character(), Instituicoes = character()))
  }

  component_blocks <- blocks %>%
    filter(BlockID %in% component_ids) %>%
    arrange(BlockID)

  component_details <- block_author_details %>%
    filter(BlockID %in% component_ids) %>%
    arrange(BlockID, AutorIndex) %>%
    mutate(
      ORCIDTokens = map(ORCIDTokens, ~ {
        if (is.null(.x) || length(.x) == 0) {
          character()
        } else {
          .x
        }
      }),
      Instituicoes = replace_na(Instituicoes, "")
    )

  if (nrow(component_details) == 0) {
    return(tibble(Ordem = min(component_blocks$BlockID), Autores = "", ORCIDs = "", Instituicoes = ""))
  }

  author_order <- component_details %>%
    distinct(Autor, .keep_all = TRUE) %>%
    mutate(AutorOrder = row_number()) %>%
    select(Autor, AutorOrder)

  author_summary_component <- component_details %>%
    group_by(Autor) %>%
    summarise(
      ORCIDValue = {
        tokens <- flatten_chr(ORCIDTokens)
        combined <- combine_orcid(tokens)
        ifelse(is.na(combined), "", combined)
      },
      InstituicoesValue = {
        combined_inst <- split_affiliations(Instituicoes)
        ifelse(is.na(combined_inst), "", combined_inst)
      },
      .groups = "drop"
    )

  author_output <- author_order %>%
    left_join(author_summary_component, by = "Autor") %>%
    mutate(
      ORCIDValue = coalesce(ORCIDValue, ""),
      InstituicoesValue = coalesce(InstituicoesValue, "")
    ) %>%
    arrange(AutorOrder)

  tibble(
    Ordem = min(component_blocks$BlockID),
    Autores = str_squish(paste(author_output$Autor, collapse = "; ")),
    ORCIDs = str_squish(paste(author_output$ORCIDValue, collapse = "; ")),
    Instituicoes = str_squish(paste(author_output$InstituicoesValue, collapse = "; "))
  )
})

if (nrow(surname_initial_groups) == 0) {
  surname_initial_groups <- tibble(
    Autores = character(),
    ORCIDs = character(),
    Instituicoes = character()
  )
} else {
  surname_initial_groups <- surname_initial_groups %>%
    arrange(Ordem) %>%
    select(Autores, ORCIDs, Instituicoes)
}

if (nrow(orcid_groups) == 0) {
  orcid_groups <- tibble(
    Autores = character(),
    ORCIDs = character(),
    Instituicoes = character()
  )
}

output_path <- file.path(folder_path, "autores_unicos.xlsx")
write_xlsx(
  list(
    "Autores" = result,
    "ORCID Agrupados" = orcid_groups,
    "Apelido Inicial Agrupados" = surname_initial_groups
  ),
  output_path
)

message("Ficheiro criado: ", output_path)
