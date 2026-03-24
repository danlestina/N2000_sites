# vibecoding s ChatGPT
# =========================
# 1) Balíèky
# =========================

install.packages(c("readxl", "dplyr", "tidyr", "writexl"))

library(readxl)
library(dplyr)
library(tidyr)
library(writexl)



# =========================
# 2) Soubory
# =========================

old_file <- "Czechia_prefill_MSsub4.xlsx"
new_file <- "Czechia_prefill_commdecis4.xlsx"


# =========================
# 3) Klíèe pro jednotlivé listy
# =========================
# Tady øíkáš, podle èeho se mají øádky na každém listu párovat.

sheet_keys <- list(
  
  SiteIdentification = c("F_1_2_site_code"),
  
  BiogeographicalMarineRegions = c("site_code", "F_2_3_1_biogeo_marine_code"),
  
  AdministrativeRegion = c("site_code", "F_2_2_1_site_nutscode"),
  
  HabitatsPresent = c("site_code", "F_3_1_1_habitat_code", "F_3_1_2_habitat_priority"),
  
  HabitatsNotPresent = c("site_code", "F_3_1_1_habitat_code_NP"),
  
  SpeciesPresent = c("site_code", "F_3_2_2_species_code", "F_3_2_6_species_poptype"),
  
  SpeciesNotPresent = c("site_code", "F_3_2_2_species_code_NP"),
  
  OtherSpecies = c("site_code", "F_3_3_1_species_group", "F_3_3_2_species_code"),
  
  SiteDescription = c("site_code"),
  
  PressuresList = c(
    "site_code",
    "F_4_3_1_pressure_code",
    "F_4_3_2_pressure_rank",
    "F_4_3_3_pressure_location"
  ),
  
  ManagementInfo = c("site_code"),
  
  ConservationMeasures_Documents = c(
    "site_code",
    "F_5_3_1_b_measures_title",
    "F_5_3_1_c_measures_URI"
  ),
  
  ManagementPlansList = c(
    "site_code",
    "F_5_2_2_a_management_plan_name"
  ),
  
  ManagementBody = c(
    "site_code",
    "F_5_1_1_a_management_body"
  )
)


# =========================
# 4) Naètení názvù listù
# =========================

old_sheets <- excel_sheets(old_file)
new_sheets <- excel_sheets(new_file)

# vezmeme sjednocení listù z obou souborù
all_sheets <- union(old_sheets, new_sheets)

all_sheets


# =========================
# 5) Pøipravení prázdných výsledkù
# =========================

summary_table <- data.frame()

all_diffs <- data.frame()
all_only_in_old <- data.frame()
all_only_in_new <- data.frame()


# =========================
# 6) Hlavní cyklus pøes všechny listy
# =========================

for (sheet_name in all_sheets) {
  
  cat("\n=============================\n")
  cat("Zpracovávám list:", sheet_name, "\n")
  cat("=============================\n")
  
  # -------------------------
  # 6a) Kontrola, že list je v obou souborech
  # -------------------------
  if (!(sheet_name %in% old_sheets)) {
    summary_table <- bind_rows(
      summary_table,
      data.frame(
        sheet = sheet_name,
        status = "missing_in_old",
        rows_old = NA,
        rows_new = NA,
        same_order = NA,
        only_in_old = NA,
        only_in_new = NA,
        changed_cells = NA
      )
    )
    next
  }
  
  if (!(sheet_name %in% new_sheets)) {
    summary_table <- bind_rows(
      summary_table,
      data.frame(
        sheet = sheet_name,
        status = "missing_in_new",
        rows_old = NA,
        rows_new = NA,
        same_order = NA,
        only_in_old = NA,
        only_in_new = NA,
        changed_cells = NA
      )
    )
    next
  }
  
  # -------------------------
  # 6b) Kontrola, že pro list máme definovaný klíè
  # -------------------------
  if (!(sheet_name %in% names(sheet_keys))) {
    summary_table <- bind_rows(
      summary_table,
      data.frame(
        sheet = sheet_name,
        status = "no_key_defined",
        rows_old = NA,
        rows_new = NA,
        same_order = NA,
        only_in_old = NA,
        only_in_new = NA,
        changed_cells = NA
      )
    )
    next
  }
  
  key_cols <- sheet_keys[[sheet_name]]
  
  # -------------------------
  # 6c) Naètení listu
  # -------------------------
  old <- read_excel(old_file, sheet = sheet_name, col_types = "text")
  new <- read_excel(new_file, sheet = sheet_name, col_types = "text")
  
  # -------------------------
  # 6d) Jednoduché vyèištìní
  # -------------------------
  # - oøezání mezer na zaèátku a na konci
  # - prázdný string -> NA
  
  old[] <- lapply(old, function(x) {
    x <- trimws(x)
    x[x == ""] <- NA
    x
  })
  
  new[] <- lapply(new, function(x) {
    x <- trimws(x)
    x[x == ""] <- NA
    x
  })
  
  # odstranìní úplnì prázdných øádkù
  old <- old[rowSums(!is.na(old)) > 0, ]
  new <- new[rowSums(!is.na(new)) > 0, ]
  
  # -------------------------
  # 6e) Jen spoleèné sloupce
  # -------------------------
  common_cols <- intersect(names(old), names(new))
  
  old <- old[, common_cols]
  new <- new[, common_cols]
  
  # -------------------------
  # 6f) Kontrola, že klíèové sloupce existují
  # -------------------------
  if (!all(key_cols %in% names(old)) || !all(key_cols %in% names(new))) {
    summary_table <- bind_rows(
      summary_table,
      data.frame(
        sheet = sheet_name,
        status = "missing_key_columns",
        rows_old = nrow(old),
        rows_new = nrow(new),
        same_order = NA,
        only_in_old = NA,
        only_in_new = NA,
        changed_cells = NA
      )
    )
    next
  }
  
  # -------------------------
  # 6g) Kontrola poøadí øádkù
  # -------------------------
  same_order <- identical(old, new)
  
  # -------------------------
  # 6h) Pøidání klíèe
  # -------------------------
  old$key <- apply(old[, key_cols, drop = FALSE], 1, paste, collapse = " | ")
  new$key <- apply(new[, key_cols, drop = FALSE], 1, paste, collapse = " | ")
  
  # -------------------------
  # 6i) Øádky jen ve starém / jen v novém
  # -------------------------
  only_in_old <- old %>%
    filter(!(key %in% new$key)) %>%
    mutate(sheet = sheet_name)
  
  only_in_new <- new %>%
    filter(!(key %in% old$key)) %>%
    mutate(sheet = sheet_name)
  
  # -------------------------
  # 6j) Spojení podle klíèe
  # -------------------------
  joined <- full_join(
    old,
    new,
    by = "key",
    suffix = c("_old", "_new")
  )
  
  # -------------------------
  # 6k) Pøevod do long formátu
  # -------------------------
  # Ze všech *_old sloupcù udìláme:
  # key | column | old_value
  #
  # Ze všech *_new sloupcù udìláme:
  # key | column | new_value
  
  old_long <- joined %>%
    select(key, ends_with("_old")) %>%
    pivot_longer(
      cols = -key,
      names_to = "column",
      values_to = "old_value"
    ) %>%
    mutate(column = sub("_old$", "", column))
  
  new_long <- joined %>%
    select(key, ends_with("_new")) %>%
    pivot_longer(
      cols = -key,
      names_to = "column",
      values_to = "new_value"
    ) %>%
    mutate(column = sub("_new$", "", column))
  
  # -------------------------
  # 6l) Porovnání bunìk
  # -------------------------
  diff_table <- full_join(old_long, new_long, by = c("key", "column")) %>%
    filter(
      !(is.na(old_value) & is.na(new_value)) &
        coalesce(old_value, "<NA>") != coalesce(new_value, "<NA>")
    ) %>%
    mutate(sheet = sheet_name) %>%
    select(sheet, key, column, old_value, new_value) %>%
    arrange(key, column)
  
  # -------------------------
  # 6m) Pøidání do celkových výsledkù
  # -------------------------
  all_diffs <- bind_rows(all_diffs, diff_table)
  all_only_in_old <- bind_rows(all_only_in_old, only_in_old)
  all_only_in_new <- bind_rows(all_only_in_new, only_in_new)
  
  # -------------------------
  # 6n) Souhrn za list
  # -------------------------
  summary_table <- bind_rows(
    summary_table,
    data.frame(
      sheet = sheet_name,
      status = "ok",
      rows_old = nrow(old),
      rows_new = nrow(new),
      same_order = same_order,
      only_in_old = nrow(only_in_old),
      only_in_new = nrow(only_in_new),
      changed_cells = nrow(diff_table)
    )
  )
}


# =========================
# 7) Výstupy v R
# =========================

summary_table
all_diffs
all_only_in_old
all_only_in_new


# =========================
# 8) Uložení do Excelu
# =========================

write_xlsx(
  list(
    summary = summary_table,
    changed_cells = all_diffs,
    only_in_old = all_only_in_old,
    only_in_new = all_only_in_new
  ),
  path = "diff_all_sheets.xlsx"
)
