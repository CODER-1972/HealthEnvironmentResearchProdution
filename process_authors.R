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

format_duration <- function(seconds) {
  if (is.na(seconds) || !is.finite(seconds)) {
    return("indisponível")
  }
  total_seconds <- as.integer(round(seconds))
  if (total_seconds < 0) {
    total_seconds <- 0
  }
  hours <- total_seconds %/% 3600
  minutes <- (total_seconds %% 3600) %/% 60
  secs <- total_seconds %% 60
  sprintf("%02d:%02d:%02d", hours, minutes, secs)
}

format_eta <- function(reference_time, seconds_remaining) {
  if (is.na(seconds_remaining) || !is.finite(seconds_remaining)) {
    return("indisponível")
  }
  eta_time <- reference_time + seconds_remaining
  format(eta_time, "%Y-%m-%d %H:%M:%S")
}

create_record_progress <- function(total_records) {
  state <- new.env(parent = emptyenv())
  state$total <- if (is.na(total_records) || total_records < 1) 0L else as.integer(total_records)
  state$current <- 0L
  state$start_time <- Sys.time()
  state$finalized <- FALSE

  state$update <- function(increment = 1L) {
    if (state$total == 0L) {
      return()
    }
    increment <- as.integer(increment)
    if (is.na(increment) || increment <= 0L) {
      increment <- 0L
    }
    state$current <- min(state$total, state$current + increment)
    now <- Sys.time()
    elapsed <- as.numeric(difftime(now, state$start_time, units = "secs"))
    average <- if (state$current == 0L) NA_real_ else elapsed / state$current
    remaining <- if (is.na(average)) NA_real_ else (state$total - state$current) * average
    eta <- format_eta(now, remaining)
    cat(
      sprintf(
        "[Registo %d/%d] Tempo decorrido: %s | Tempo restante estimado: %s | Conclusão estimada às %s\n",
        state$current,
        state$total,
        format_duration(elapsed),
        format_duration(remaining),
        eta
      )
    )
    flush.console()
  }

  state$finish <- function() {
    if (state$finalized) {
      return()
    }
    state$finalized <- TRUE
    now <- Sys.time()
    elapsed <- as.numeric(difftime(now, state$start_time, units = "secs"))
    cat(
      sprintf(
        "Processamento concluído. Tempo total decorrido: %s para %d registos.\n",
        format_duration(elapsed),
        state$total
      )
    )
    flush.console()
  }

  state
}

create_step_tracker <- function(step_names) {
  state <- new.env(parent = emptyenv())
  state$step_names <- if (is.null(step_names)) character() else as.character(step_names)
  state$total <- length(state$step_names)
  state$current <- 0L
  state$durations <- numeric()
  state$step_start <- NULL
  state$active_name <- NULL
  state$finished <- FALSE

  state$begin <- function(step_name = NULL) {
    if (state$finished) {
      return()
    }
    next_index <- state$current + 1L
    if (!is.null(step_name)) {
      name <- as.character(step_name)[1]
    } else if (next_index <= length(state$step_names)) {
      name <- state$step_names[[next_index]]
    } else {
      name <- sprintf("Passo %d", next_index)
    }
    estimate <- if (length(state$durations) == 0) {
      NA_real_
    } else {
      mean(tail(state$durations, n = min(length(state$durations), 5L)))
    }
    total_steps <- max(length(state$step_names), next_index)
    cat(
      sprintf(
        "[Passo %d/%d] A iniciar \"%s\" | Duração estimada deste passo: %s\n",
        next_index,
        total_steps,
        name,
        format_duration(estimate)
      )
    )
    flush.console()
    state$current <- next_index
    state$active_name <- name
    state$step_start <- Sys.time()
    invisible(NULL)
  }

  state$end <- function(success = TRUE) {
    if (state$finished || is.null(state$step_start)) {
      return()
    }
    now <- Sys.time()
    duration <- as.numeric(difftime(now, state$step_start, units = "secs"))
    state$durations <- c(state$durations, duration)
    status_text <- if (isTRUE(success)) "concluído" else "terminado com erro"
    total_steps <- max(length(state$step_names), state$current)
    active_name <- if (is.null(state$active_name)) "" else state$active_name
    cat(
      sprintf(
        "[Passo %d/%d] \"%s\" %s em %s | Passos concluídos: %d/%d\n",
        state$current,
        total_steps,
        active_name,
        status_text,
        format_duration(duration),
        state$current,
        total_steps
      )
    )
    flush.console()
    state$step_start <- NULL
    state$active_name <- NULL
    invisible(NULL)
  }

  state$finish <- function() {
    if (state$finished) {
      return()
    }
    if (!is.null(state$step_start)) {
      state$end(success = FALSE)
    }
    total_steps <- max(length(state$step_names), state$current)
    total_duration <- if (length(state$durations) == 0) 0 else sum(state$durations)
    cat(
      sprintf(
        "Resumo dos passos: %d/%d concluídos | Duração total aproximada: %s\n",
        state$current,
        total_steps,
        format_duration(total_duration)
      )
    )
    flush.console()
    state$finished <- TRUE
    invisible(NULL)
  }

  state
}
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

total_records <- nrow(data)
record_progress <- create_record_progress(total_records)
on.exit(record_progress$finish(), add = TRUE)
cat(sprintf("Total de registos a processar: %d\n", total_records))

step_plan <- c(
  "Detetar colunas principais",
  "Extrair dados por registo",
  "Conciliar identificadores ORCID",
  "Associar instituições aos autores",
  "Construir resumo principal de autores",
  "Construir agrupamento por ORCID",
  "Preparar blocos e tokens de instituições",
  "Identificar grupos por apelido/inicial",
  "Gerar tabelas de agrupamento de instituições",
  "Exportar ficheiros"
)
step_tracker <- create_step_tracker(step_plan)
on.exit(step_tracker$finish(), add = TRUE)

with_step <- function(step_name, expr) {
  step_tracker$begin(step_name)
  result <- tryCatch(
    eval.parent(substitute(expr)),
    error = function(e) {
      step_tracker$end(success = FALSE)
      stop(e)
    }
  )
  step_tracker$end(success = TRUE)
  result
}

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

`%||%` <- function(x, y) {
  if (is.null(x) || length(x) == 0) {
    y
  } else {
    x
  }
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
  if (is.null(value)) {
    return(character())
  }

  value <- as.character(value)
  if (length(value) == 0) {
    return(value)
  }

  value[is.na(value)] <- ""

  stopwords <- c(
    "de", "da", "do", "das", "dos", "os", "as", "e", "and", "the", "of",
    "del", "della", "di", "du", "des", "el", "la", "le", "los",
    "las", "en", "y", "a", "an",
    "portugal", "germany", "spain", "italy", "france", "europe", "portuguese",
    "german", "spanish", "italian", "french", "avenue", "street", "road",
    "campus", "building", "city", "town", "ulmenliet", "aveiro", "faro",
    "hamburg", "lisbon", "porto"
  )

  normalize_word <- function(word) {
    if (word == "") {
      return("")
    }

    if (word %in% stopwords) {
      return("")
    }

    canonical <- word
    canonical <- str_replace(canonical, "^(univ|univers[a-z]+)$", "university")
    canonical <- str_replace(canonical, "^(dept|depto|dep[a-z]*|depart[a-z]*)$", "department")
    canonical <- str_replace(canonical, "^(ctr|cent[roes]*|centre[s]?)$", "center")
    canonical <- str_replace(canonical, "^(inst|institu[a-z]*|institution[a-z]*)$", "institute")
    canonical <- str_replace(canonical, "^(tech|technol[a-z]*|tecnolog[a-z]*)$", "technology")
    canonical <- str_replace(canonical, "^(sci|science[s]?|scient[a-z]*)$", "science")
    canonical <- str_replace(canonical, "^(res|research[a-z]*|recherche)$", "research")
    canonical <- str_replace(canonical, "^(agro[a-z]*|agr[a-z]*|agriculture[a-z]*)$", "agriculture")
    canonical <- str_replace(canonical, "^(environ[a-z]*|ambient[a-z]*)$", "environment")
    canonical <- str_replace(canonical, "^(biol[a-z]*|bio[a-z]*)$", "biology")
    canonical <- str_replace(canonical, "^(hlth|health|saude)$", "health")
    canonical <- str_replace(canonical, "^(sport[a-z]*|desport[a-z]*)$", "sport")
    canonical <- str_replace(canonical, "^(lab|labor[a-z]*)$", "laboratory")
    canonical <- str_replace(canonical, "^(fac|faculd[a-z]*|facult[a-z]*)$", "faculty")
    canonical <- str_replace(canonical, "^(vilareal|vila?real)$", "vilareal")
    canonical <- str_replace(canonical, "^(tras|tros)$", "trasosmontes")
    canonical <- str_replace(canonical, "^(altodouro|alto)$", "altodouro")
    canonical <- str_replace(canonical, "^(citab)$", "citab")

    canonical <- str_replace_all(canonical, "(.)\\1+", "\\1")

    if (str_detect(canonical, "[0-9]")) {
      return("")
    }

    if (nchar(canonical) <= 2) {
      return("")
    }

    if (nchar(canonical) >= 3) {
      consonant_signature <- str_replace_all(canonical, "[aeiou]", "")
      if (nchar(consonant_signature) >= 2) {
        canonical <- consonant_signature
      }
    }

    if (nchar(canonical) > 10) {
      canonical <- str_sub(canonical, 1, 10)
    }

    canonical
  }

  normalized <- strip_diacritics(value)
  normalized <- str_to_lower(normalized)
  normalized <- str_replace_all(normalized, "[&/@]", " ")
  normalized <- str_replace_all(normalized, "tras\\s+os?\\s+montes", "trasosmontes")
  normalized <- str_replace_all(normalized, "tros\\s+montes", "trasosmontes")
  normalized <- str_replace_all(normalized, "alto\\s+douro", "altodouro")
  normalized <- str_replace_all(normalized, "vila\\s+real", "vilareal")
  normalized <- str_replace_all(normalized, "agro\\s+environ", "agroenviron")
  normalized <- str_replace_all(normalized, "agroenvironm", "agroenviron")
  normalized <- str_replace_all(normalized, "[^a-z0-9]+", " ")
  normalized <- str_squish(normalized)

  vapply(
    normalized,
    function(entry) {
      if (entry == "") {
        return("")
      }

      words <- str_split(entry, "\\s+")[[1]]
      words <- words[words != ""]
      if (length(words) == 0) {
        return("")
      }

      canonical_words <- vapply(words, normalize_word, character(1))
      canonical_words <- canonical_words[canonical_words != ""]
      if (length(canonical_words) == 0) {
        return("")
      }

      canonical_words <- sort(unique(canonical_words))
      paste(canonical_words, collapse = " ")
    },
    character(1),
    USE.NAMES = FALSE
  )
}

institution_similarity_score <- function(a, b) {
  if (length(a) == 0 || length(b) == 0) {
    return(0)
  }

  if (a == "" || b == "") {
    return(0)
  }

  tokens_a <- unique(str_split(a, "\\s+")[[1]])
  tokens_b <- unique(str_split(b, "\\s+")[[1]])

  tokens_a <- tokens_a[tokens_a != ""]
  tokens_b <- tokens_b[tokens_b != ""]

  filtered_a <- tokens_a[!str_detect(tokens_a, "[0-9]") & nchar(tokens_a) > 2]
  filtered_b <- tokens_b[!str_detect(tokens_b, "[0-9]") & nchar(tokens_b) > 2]

  calc_jaccard <- function(x, y) {
    if (length(x) == 0 || length(y) == 0) {
      return(0)
    }
    intersection <- length(intersect(x, y))
    union <- length(union(x, y))
    if (union == 0) 0 else intersection / union
  }

  jaccard_full <- calc_jaccard(tokens_a, tokens_b)
  jaccard_filtered <- calc_jaccard(filtered_a, filtered_b)
  jaccard <- max(jaccard_full, jaccard_filtered)

  max_len <- max(nchar(a), nchar(b), 1)
  distance <- adist(a, b)
  char_similarity <- 1 - (distance / max_len)

  max(jaccard, char_similarity)
}

compute_institution_similarity <- function(keys) {
  keys <- unique(keys[keys != ""])
  if (length(keys) < 2) {
    return(tibble(Key1 = character(), Key2 = character(), Score = numeric()))
  }

  total_pairs <- choose(length(keys), 2)
  if (total_pairs == 0) {
    return(tibble(Key1 = character(), Key2 = character(), Score = numeric()))
  }

  results <- vector("list", total_pairs)
  edge_index <- 1L

  for (i in seq_len(length(keys) - 1)) {
    for (j in seq((i + 1), length(keys))) {
      score <- institution_similarity_score(keys[[i]], keys[[j]])
      if (!is.na(score) && score > 0) {
        results[[edge_index]] <- tibble(
          Key1 = keys[[i]],
          Key2 = keys[[j]],
          Score = score
        )
        edge_index <- edge_index + 1L
      }
    }
  }

  if (edge_index == 1L) {
    tibble(Key1 = character(), Key2 = character(), Score = numeric())
  } else {
    bind_rows(results[seq_len(edge_index - 1L)])
  }
}

build_clusters_from_edges <- function(keys, edges) {
  if (length(keys) == 0) {
    return(list())
  }

  key_index <- seq_along(keys)
  names(key_index) <- keys
  parent <- seq_along(keys)

  find_parent <- function(idx) {
    while (parent[[idx]] != idx) {
      parent[[idx]] <<- parent[[parent[[idx]]]]
      idx <- parent[[idx]]
    }
    idx
  }

  union_parent <- function(key_a, key_b) {
    idx_a <- key_index[[key_a]]
    idx_b <- key_index[[key_b]]
    if (is.na(idx_a) || is.na(idx_b)) {
      return()
    }
    root_a <- find_parent(idx_a)
    root_b <- find_parent(idx_b)
    if (root_a == root_b) {
      return()
    }
    parent[[root_b]] <<- root_a
  }

  if (nrow(edges) > 0) {
    for (i in seq_len(nrow(edges))) {
      union_parent(edges$Key1[[i]], edges$Key2[[i]])
    }
  }

  groups <- vapply(seq_along(keys), find_parent, integer(1))
  split(keys, groups)
}

cluster_institution_keys <- function(keys, threshold = 0.80, edges = NULL) {
  unique_keys <- unique(keys[keys != ""])
  if (length(unique_keys) == 0) {
    return(tibble(InstituicaoKey = character(), ClusterKey = character()))
  }
  if (length(unique_keys) == 1) {
    return(tibble(InstituicaoKey = unique_keys, ClusterKey = unique_keys))
  }

  if (is.null(edges)) {
    similarity_edges <- compute_institution_similarity(unique_keys)
  } else {
    similarity_edges <- edges %>%
      filter(Key1 %in% unique_keys & Key2 %in% unique_keys)
  }

  if (nrow(similarity_edges) == 0) {
    return(tibble(InstituicaoKey = unique_keys, ClusterKey = unique_keys))
  }

  relevant_edges <- similarity_edges %>%
    filter(Score >= threshold)

  if (nrow(relevant_edges) == 0) {
    return(tibble(InstituicaoKey = unique_keys, ClusterKey = unique_keys))
  }

  all_keys <- sort(unique(unique_keys))
  clusters <- build_clusters_from_edges(all_keys, relevant_edges)

  cluster_lookup <- setNames(all_keys, all_keys)
  for (members in clusters) {
    representative <- members[which.min(nchar(members))]
    cluster_lookup[members] <- representative
  }

  tibble(
    InstituicaoKey = unique_keys,
    ClusterKey = unname(cluster_lookup[unique_keys])
  )
}

cluster_keys_in_range <- function(edges, lower, upper, excluded_keys = character()) {
  if (is.null(edges) || nrow(edges) == 0) {
    return(list())
  }

  range_edges <- edges %>%
    filter(Score >= lower & Score < upper)

  if (length(excluded_keys) > 0) {
    range_edges <- range_edges %>%
      filter(!(Key1 %in% excluded_keys | Key2 %in% excluded_keys))
  }

  if (nrow(range_edges) == 0) {
    return(list())
  }

  keys_range <- sort(unique(c(range_edges$Key1, range_edges$Key2)))
  build_clusters_from_edges(keys_range, range_edges)
}

choose_standard_name <- function(variants, fallback = "") {
  values <- variants[!is.na(variants) & variants != ""]
  if (length(values) == 0) {
    if (is.null(fallback) || fallback == "") {
      return("")
    }
    return(fallback)
  }

  ordered <- values[order(nchar(values), values)]
  ordered[[1]]
}

build_cluster_table <- function(clusters, tokens_lookup) {
  valid_clusters <- clusters[sapply(clusters, length) > 1]
  if (length(valid_clusters) == 0) {
    return(tibble(`Nome standard` = character(), `Variantes agrupadas` = character()))
  }

  map_dfr(valid_clusters, function(member_keys) {
    variant_lists <- map(member_keys, function(key) {
      value <- tokens_lookup[[key]]
      if (is.null(value) || length(value) == 0) {
        key
      } else {
        value
      }
    })

    variant_values <- variant_lists %>%
      map(~ as.character(.x)) %>%
      flatten_chr() %>%
      str_trim()

    if (length(variant_values) == 0) {
      variant_values <- member_keys
    }

    unique_variants <- unique(variant_values[variant_values != ""])
    standard <- choose_standard_name(unique_variants, fallback = member_keys[[1]])
    standard <- str_squish(standard)
    aggregated <- unique(c(standard, unique_variants))
    aggregated <- aggregated[aggregated != ""]
    aggregated_str <- if (length(aggregated) == 0) "" else paste(aggregated, collapse = "; ")

    tibble(
      `Nome standard` = standard,
      `Variantes agrupadas` = aggregated_str
    )
  }) %>%
    arrange(`Nome standard`)
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

with_step("Detetar colunas principais", {
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

  author_raw <- data[[full_name_col]]
})

with_step("Extrair dados por registo", {
  author_entries_list <<- vector("list", total_records)
  orcid_entries_list <<- vector("list", total_records)
  affiliation_entries_list <<- vector("list", total_records)

  for (row_id in seq_len(total_records)) {
    current_authors <- split_authors(author_raw[[row_id]])
    if (length(current_authors) > 0) {
      author_entries <- tibble(RowID = row_id, Autor = current_authors) %>%
        mutate(
          Autor = str_squish(Autor),
          AutorKey = make_author_key(Autor)
        ) %>%
        filter(!is.na(Autor) & Autor != "")
    } else {
      author_entries <- tibble(RowID = integer(), Autor = character(), AutorKey = character())
    }
    author_entries_list[[row_id]] <- author_entries

    if (!is.null(orcid_col)) {
      current_orcid <- parse_orcid_entries(data[[orcid_col]][[row_id]])
      if (nrow(current_orcid) > 0) {
        current_orcid <- current_orcid %>%
          mutate(RowID = row_id) %>%
          select(RowID, AutorKey, ORCID)
      } else {
        current_orcid <- tibble(RowID = integer(), AutorKey = character(), ORCID = character())
      }
    } else {
      current_orcid <- tibble(RowID = integer(), AutorKey = character(), ORCID = character())
    }
    orcid_entries_list[[row_id]] <- current_orcid

    if (!is.null(affiliation_col)) {
      current_affiliation <- parse_affiliation_entries(data[[affiliation_col]][[row_id]], row_id)
    } else {
      current_affiliation <- tibble(RowID = integer(), AutorKey = character(), Affiliation = character())
    }
    affiliation_entries_list[[row_id]] <- current_affiliation

    record_progress$update()
  }

  # Garante que o utilizador vê de imediato que a leitura linha a linha terminou
  # antes de avançar para as fases de consolidação que podem ser mais demoradas.
  record_progress$finish()
  cat(
    "Leitura de registos concluída. A consolidar ORCID e instituições; esta etapa pode demorar alguns instantes...\n"
  )
  flush.console()

  author_rows <<- if (length(author_entries_list) == 0) {
    tibble(RowID = integer(), Autor = character(), AutorKey = character())
  } else {
    bind_rows(author_entries_list)
  }

  if (nrow(author_rows) == 0) {
    stop("Não foram encontrados autores no ficheiro fornecido.")
  }

  author_counts <- author_rows %>%
    count(RowID, name = "AutorCount")

  author_rows <<- author_rows %>%
    left_join(author_counts, by = "RowID")
})

with_step("Conciliar identificadores ORCID", {
  if (is.null(orcid_col)) {
    message("Aviso: coluna de ORCID não encontrada. Será criado um campo vazio.")
  }

  orcid_map <- if (!is.null(orcid_col)) {
    if (length(orcid_entries_list) == 0) {
      tibble(RowID = integer(), AutorKey = character(), ORCID = character())
    } else {
      bind_rows(orcid_entries_list)
    }
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

  authors_with_orcid <<- author_rows %>%
    left_join(orcid_specific, by = c("RowID", "AutorKey"))

  if (nrow(orcid_unspecified) > 0) {
    authors_with_orcid <<- authors_with_orcid %>%
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

  orcid_orphans <<- orcid_map %>%
    filter((is.na(AutorKey) | AutorKey == "") & !is.na(ORCID) & ORCID != "") %>%
    distinct(RowID, ORCID) %>%
    anti_join(assigned_orcids, by = c("RowID", "ORCID")) %>%
    distinct(ORCID)
})


with_step("Associar instituições aos autores", {
  if (is.null(affiliation_col)) {
    message("Aviso: coluna de filiação/instituições não encontrada. Será criado um campo vazio.")
    affiliation_map <- tibble(RowID = integer(), AutorKey = character(), Affiliation = character())
  } else {
    if (length(affiliation_entries_list) == 0) {
      affiliation_map <- tibble(RowID = integer(), AutorKey = character(), Affiliation = character())
    } else {
      affiliation_map <- bind_rows(affiliation_entries_list)
    }
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

  authors_combined <<- authors_with_orcid %>%
    left_join(affiliation_combined, by = c("RowID", "AutorKey"))
})


with_step("Construir resumo principal de autores", {
  author_summary <<- authors_combined %>%
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

  author_summary <<- author_summary %>%
    mutate(OrderIndex = row_number())
  author_summary <<- author_summary %>%
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
    author_summary <<- bind_rows(author_summary, orphan_rows)
  }

  author_summary <<- author_summary %>%
    arrange(OrderIndex)

  result <<- author_summary %>%
    select(Autor, ORCID, Instituicoes) %>%
    mutate(
      AutorOrdenacao = str_trim(coalesce(Autor, ""))
    ) %>%
    arrange(AutorOrdenacao == "", AutorOrdenacao) %>%
    select(-AutorOrdenacao)
})


with_step("Construir agrupamento por ORCID", {
  author_rows_for_groups <<- author_summary %>%
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

  with_orcid <<- author_rows_for_groups %>%
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

  orcid_groups_enriched <<- bind_rows(grouped_orcid, without_orcid) %>%
    mutate(
      PrimeiroAutor = str_trim(coalesce(str_split_fixed(Autores, ";", 2)[, 1], ""))
    ) %>%
    arrange(PrimeiroAutor == "", PrimeiroAutor, Autores) %>%
    select(-PrimeiroAutor)

  orcid_groups <<- orcid_groups_enriched %>%
    select(Autores, ORCIDs, Instituicoes)
})


with_step("Preparar blocos e tokens de instituições", {
  blocks <<- orcid_groups_enriched %>%
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

  block_author_details <<- blocks %>%
    select(BlockID, Detalhes) %>%
    unnest(Detalhes)

  if (nrow(block_author_details) == 0) {
    block_author_details <<- tibble(
      BlockID = integer(),
      Autor = character(),
      ORCIDTokens = list(),
      Instituicoes = character(),
      AutorIndex = integer()
    )
  }

  block_author_details <<- block_author_details %>%
    mutate(
      SurnameInitial = map_chr(Autor, extract_surname_initial),
      InstituicaoTokens = map(Instituicoes, collect_affiliation_tokens)
    )

  institution_key_tokens <<- block_author_details %>%
    select(InstituicaoTokens) %>%
    mutate(
      InstituicaoTokens = map(InstituicaoTokens, ~ .x[!is.na(.x) & .x != ""])
    ) %>%
    filter(map_int(InstituicaoTokens, length) > 0) %>%
    unnest_longer(InstituicaoTokens, values_to = "InstituicaoToken") %>%
    mutate(
      InstituicaoKey = normalize_institution_token(InstituicaoToken)
    ) %>%
    filter(InstituicaoKey != "") %>%
    group_by(InstituicaoKey) %>%
    summarise(
      Variantes = list(sort(unique(InstituicaoToken))),
      .groups = "drop"
    )

  author_similarity_tokens <<- block_author_details %>%
    filter(!is.na(SurnameInitial) & SurnameInitial != "") %>%
    mutate(InstituicaoTokens = map(InstituicaoTokens, ~ unique(.x))) %>%
    filter(map_int(InstituicaoTokens, length) > 0) %>%
    unnest_longer(InstituicaoTokens, values_to = "InstituicaoToken") %>%
    mutate(
      InstituicaoKey = normalize_institution_token(InstituicaoToken)
    ) %>%
    filter(InstituicaoKey != "") %>%
    distinct(BlockID, SurnameInitial, InstituicaoKey)
})

with_step("Identificar grupos por apelido/inicial", {
  institution_similarity_edges <<- tibble(Key1 = character(), Key2 = character(), Score = numeric())
  institution_clusters <<- tibble(InstituicaoKey = character(), ClusterKey = character())

  if (nrow(author_similarity_tokens) > 0) {
    institution_similarity_edges <<- compute_institution_similarity(author_similarity_tokens$InstituicaoKey)
    institution_clusters <<- cluster_institution_keys(
      author_similarity_tokens$InstituicaoKey,
      threshold = 0.80,
      edges = institution_similarity_edges
    )
    author_similarity_tokens <<- author_similarity_tokens %>%
      left_join(institution_clusters, by = "InstituicaoKey") %>%
      mutate(
        ClusterKey = if_else(is.na(ClusterKey) | ClusterKey == "", InstituicaoKey, ClusterKey),
        InstituicaoKey = ClusterKey
      ) %>%
      select(-ClusterKey) %>%
      distinct(BlockID, SurnameInitial, InstituicaoKey)
  }

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

  surname_initial_groups <<- map_dfr(components_blocks, function(component_ids) {
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
    surname_initial_groups <<- tibble(
      Autores = character(),
      ORCIDs = character(),
      Instituicoes = character()
    )
  } else {
    surname_initial_groups <<- surname_initial_groups %>%
      arrange(Ordem) %>%
      select(Autores, ORCIDs, Instituicoes)
  }
})


if (nrow(orcid_groups) == 0) {
  orcid_groups <- tibble(
    Autores = character(),
    ORCIDs = character(),
    Instituicoes = character()
  )
}

with_step("Gerar tabelas de agrupamento de instituições", {
  tokens_lookup <- institution_key_tokens$Variantes
  if (length(tokens_lookup) == 0) {
    tokens_lookup <- list()
  }
  names(tokens_lookup) <- institution_key_tokens$InstituicaoKey

  clusters_strong <- institution_clusters %>%
    group_by(ClusterKey) %>%
    summarise(Members = list(sort(unique(InstituicaoKey))), .groups = "drop") %>%
    pull(Members)

  group_table_80 <<- build_cluster_table(clusters_strong, tokens_lookup)

  strong_member_keys <- if (length(clusters_strong) == 0) {
    character()
  } else {
    unique(unlist(clusters_strong[sapply(clusters_strong, length) > 1]))
  }

  clusters_mid <- cluster_keys_in_range(
    institution_similarity_edges,
    lower = 0.70,
    upper = 0.80,
    excluded_keys = strong_member_keys
  )

  group_table_70_80 <<- build_cluster_table(clusters_mid, tokens_lookup)
})

with_step("Exportar ficheiros", {
  aggregation_output_path <- file.path(folder_path, "agrupamento_instituicoes.xlsx")

  write_xlsx(
    list(
      "Similaridade >= 0.80" = group_table_80,
      "Similaridade 0.70-0.80" = group_table_70_80
    ),
    aggregation_output_path
  )

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
  message("Ficheiro de agrupamentos criado: ", aggregation_output_path)
})
