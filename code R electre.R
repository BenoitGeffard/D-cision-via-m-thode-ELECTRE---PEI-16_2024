# Charger les bibliothèques nécessaires
library(tidyverse)
library(readxl)

# Charger le fichier Excel
data <- read_excel("D:/IA Montpellier/PEI/méthode électre/Tableau_paramètres.xlsx")

# Séparer les poids et les alternatives
Pj <- as.numeric(data[1, 2:11]) # Poids Pj
Pj_prime <- as.numeric(data[2, 2:11]) # Poids P'j
alternatives <- data[3:nrow(data), ]

# Convertir les colonnes C1 à C10 en numériques (gérer les cas spécifiques si nécessaire)
colnames(alternatives)[1] <- "Alternative"; alternatives[, 2:11] <- lapply(alternatives[, 2:11], as.numeric)

# Créer un dossier pour stocker les fichiers CSV
output_folder <- "ELECTRE_results"
if (!dir.exists(output_folder)) {
  dir.create(output_folder)
}

# Fonction pour créer une matrice de concordance en favorisant les petites valeurs
calculate_concordance_min <- function(alternatives, weights) {
  n <- nrow(alternatives)
  concordance <- matrix(0, n, n)

  for (i in 1:n) {
    for (j in 1:n) {
      if (i != j) {
        concordance[i, j] <- sum(weights[alternatives[i, 2:11] <= alternatives[j, 2:11]])
      }
    }
  }
  return(concordance)
}

# Fonction pour créer une matrice de discordance en favorisant les petites valeurs
calculate_discordance_min <- function(alternatives) {
  n <- nrow(alternatives)
  discordance <- matrix(0, n, n)

  for (i in 1:n) {
    for (j in 1:n) {
      if (i != j) {
        discordance[i, j] <- max((alternatives[i, 2:11] - alternatives[j, 2:11]) /
                                   (apply(alternatives[, 2:11], 2, max) - apply(alternatives[, 2:11], 2, min)))
      }
    }
  }
  return(discordance)
}

# Appliquer les seuils pour générer une matrice de surclassement
generate_ranking <- function(concordance, discordance, concordance_threshold, discordance_threshold, alternatives_info_path) {
  # Charger les informations des alternatives depuis un fichier
  alternatives_info <- read_excel(alternatives_info_path)
  
  n <- nrow(concordance)
  outranking <- matrix(0, n, n)
  
  for (i in 1:n) {
    for (j in 1:n) {
      if (i != j) {
        outranking[i, j] <- ifelse(concordance[i, j] >= concordance_threshold & 
                                     discordance[i, j] <= discordance_threshold, 1, 0)
      }
    }
  }
  
  # Calculer les scores de dominance (somme des lignes de la matrice de surclassement)
  scores <- rowSums(outranking)
  
  # Créer un dataframe pour le classement
  ranking <- data.frame(
    Alternative = alternatives[, 1],
    Score = scores
  ) %>%
    left_join(alternatives_info, by = c("Alternative" = "ai")) %>%  # Associer les infos des parcs et CRC
    arrange(desc(Score))
  
  return(list(outranking_matrix = outranking, ranking = ranking))
}

# Calcul des matrices pour Pj avec minimisation
concordance_Pj_min <- calculate_concordance_min(alternatives, Pj)
discordance_Pj_min <- calculate_discordance_min(alternatives)

# Calcul des matrices pour P'j avec minimisation
concordance_Pj_prime_min <- calculate_concordance_min(alternatives, Pj_prime)
discordance_Pj_prime_min <- calculate_discordance_min(alternatives)

# Chemin vers le fichier contenant les informations des alternatives
alternatives_info_path <- "D:/IA Montpellier/PEI/méthode électre/Noms_parcs.xlsx"

# Générer le classement avec les seuils donnés
results <- generate_ranking(concordance_Pj_min, discordance_Pj_min, 0.4, 0.9, alternatives_info_path)
results_prime <- generate_ranking(concordance_Pj_prime_min, discordance_Pj_prime_min, 0.4, 0.9, alternatives_info_path)
outranking_matrix <- results$outranking_matrix
outranking_matrix_prime <- results_prime$outranking_matrix
ranking <- results$ranking
ranking_prime <- results_prime$ranking

# Sauvegarder les résultats dans des fichiers CSV dans le dossier spécifié
write.csv(concordance_Pj_min, file.path(output_folder, "concordance_Pj_min.csv"), row.names = FALSE)
write.csv(discordance_Pj_min, file.path(output_folder, "discordance_Pj_min.csv"), row.names = FALSE)
write.csv(concordance_Pj_prime_min, file.path(output_folder, "concordance_Pj_prime_min.csv"), row.names = FALSE)
write.csv(discordance_Pj_prime_min, file.path(output_folder, "discordance_Pj_prime_min.csv"), row.names = FALSE)
write.csv(outranking_matrix, file.path(output_folder, "outranking_matrix.csv"), row.names = FALSE)
write.csv(outranking_matrix_prime, file.path(output_folder, "outranking_matrix_prime.csv"), row.names = FALSE)
write.csv(ranking, file.path(output_folder, "ranking.csv"), row.names = FALSE)
write.csv(ranking_prime, file.path(output_folder, "ranking_prime.csv"), row.names = FALSE)

# Message de confirmation
cat("Les fichiers CSV ont été sauvegardés dans le dossier :", output_folder, "\n")
cat("Les classements des alternatives ont été sauvegardés dans le dossier :", output_folder, "\n")
