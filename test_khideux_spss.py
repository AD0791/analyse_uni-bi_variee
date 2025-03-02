from docx import Document
from docx.shared import Inches

# Créer un nouveau document Word
doc = Document()

# Ajouter un titre principal
title = doc.add_heading('Guide pour effectuer un test du Khi-deux sur SPSS', level=1)
title.alignment = 1  # Centrer le titre

# Ajouter un paragraphe d'introduction
doc.add_paragraph(
    "Ce document explique comment effectuer un test du Khi-deux (χ²) sur SPSS pour déterminer s'il existe une association significative entre deux variables catégorielles. "
    "Il fournit des instructions étape par étape, des conseils pour interpréter les résultats, et des exemples concrets dans le cadre de votre étude sur l'impact des crédits agricoles sur les agriculteurs de Thomassique."
)

# Ajouter un sous-titre pour la préparation des données
doc.add_heading('1. Préparer vos données', level=2)

# Explication de la préparation des données
doc.add_paragraph(
    "Avant de commencer, assurez-vous que vos données sont correctement structurées :\n"
    "- Chaque ligne doit représenter une observation (par exemple, un agriculteur).\n"
    "- Chaque colonne doit représenter une variable (par exemple, accès au crédit, amélioration des revenus).\n"
    "- Les deux variables que vous souhaitez tester doivent être catégorielles (par exemple, Oui/Non)."
)

doc.add_paragraph(
    "Exemple de données :\n"
    "- Variable 1 : Accès au crédit (Oui/Non).\n"
    "- Variable 2 : Amélioration des revenus (Oui/Non)."
)

# Ajouter un sous-titre pour l'importation des données dans SPSS
doc.add_heading('2. Ouvrir SPSS et importer les données', level=2)

# Explication de l'importation des données
doc.add_paragraph(
    "1. Lancez SPSS sur votre ordinateur.\n"
    "2. Importez vos données :\n"
    "   - Allez dans `File > Open > Data`.\n"
    "   - Sélectionnez votre fichier de données (par exemple, un fichier Excel).\n"
    "3. Vérifiez que vos données sont correctement structurées dans SPSS."
)

# Ajouter un sous-titre pour la création d'un tableau croisé
doc.add_heading('3. Créer un tableau croisé', level=2)

# Explication de la création d'un tableau croisé
doc.add_paragraph(
    "1. Accédez à l'option Tableau croisé :\n"
    "   - Allez dans `Analyze > Descriptive Statistics > Crosstabs`.\n"
    "2. Sélectionnez les variables :\n"
    "   - Placez la première variable catégorielle (par exemple, Accès au crédit) dans la zone **Rows** (lignes).\n"
    "   - Placez la deuxième variable catégorielle (par exemple, Amélioration des revenus) dans la zone **Columns** (colonnes).\n"
    "3. Options supplémentaires :\n"
    "   - Cochez la case `Display clustered bar charts` si vous souhaitez afficher un graphique en barres."
)

# Ajouter un sous-titre pour la configuration du test du Khi-deux
doc.add_heading('4. Configurer le test du Khi-deux', level=2)

# Explication de la configuration du test
doc.add_paragraph(
    "1. Accédez aux statistiques :\n"
    "   - Dans la boîte de dialogue `Crosstabs`, cliquez sur le bouton `Statistics`.\n"
    "2. Sélectionnez le test du Khi-deux :\n"
    "   - Cochez la case `Chi-square` dans la section `Nominal`.\n"
    "   - Vous pouvez également cocher `Phi and Cramer's V` pour mesurer la force de l'association.\n"
    "3. Valider :\n"
    "   - Cliquez sur `Continue` pour revenir à la boîte de dialogue principale."
)

# Ajouter un sous-titre pour l'exécution de l'analyse
doc.add_heading("5. Lancer l'analyse", level=2)

# Explication de l'exécution de l'analyse
doc.add_paragraph(
    "1. Exécutez le test :\n"
    "   - Cliquez sur `OK` dans la boîte de dialogue `Crosstabs` pour lancer l'analyse.\n"
    "2. Résultats :\n"
    "   - SPSS générera deux tableaux principaux :\n"
    "     - **Tableau croisé** : Affiche les effectifs observés et les effectifs attendus.\n"
    "     - **Test du Khi-deux** : Affiche la statistique du Khi-deux, les degrés de liberté et la p-value."
)

# Ajouter un sous-titre pour l'interprétation des résultats
doc.add_heading('6. Interpréter les résultats', level=2)

# Explication de l'interprétation des résultats
doc.add_paragraph(
    "1. Tableau croisé :\n"
    "   - Ce tableau montre les effectifs observés pour chaque combinaison de catégories des deux variables.\n"
    "   - Exemple :\n"
    "     - Nombre d'agriculteurs ayant accès au crédit et ayant vu leurs revenus augmenter.\n"
    "     - Nombre d'agriculteurs n'ayant pas accès au crédit et n'ayant pas vu leurs revenus augmenter.\n"
    "2. Test du Khi-deux :\n"
    "   - **Statistique du Khi-deux (χ²)** : Mesure l'écart entre les effectifs observés et les effectifs attendus.\n"
    "   - **Degrés de liberté (df)** : Calculés comme (nombre de lignes - 1) × (nombre de colonnes - 1).\n"
    "   - **p-value** : Indique la probabilité d'observer les résultats (ou des résultats plus extrêmes) si l'hypothèse nulle est vraie.\n"
    "3. Interprétation :\n"
    "   - **Si la p-value est inférieure à 0.05** :\n"
    "     - On rejette l'hypothèse nulle (H₀).\n"
    "     - On conclut qu'il existe une association significative entre les deux variables.\n"
    "   - **Si la p-value est supérieure à 0.05** :\n"
    "     - On ne peut pas rejeter l'hypothèse nulle (H₀).\n"
    "     - On conclut qu'il n'y a pas d'association significative entre les deux variables."
)

# Ajouter un sous-titre pour l'exportation des résultats
doc.add_heading('7. Exporter les résultats', level=2)

# Explication de l'exportation des résultats
doc.add_paragraph(
    "1. Exporter les tableaux :\n"
    "   - Cliquez avec le bouton droit sur le tableau ou le graphique dans la fenêtre des résultats.\n"
    "   - Sélectionnez `Export` pour enregistrer les résultats dans un fichier (par exemple, Excel ou PDF).\n"
    "2. Enregistrer le fichier SPSS :\n"
    "   - Allez dans `File > Save As` pour enregistrer votre fichier de données SPSS (.sav)."
)

# Ajouter un sous-titre pour un exemple concret
doc.add_heading('Exemple concret dans votre étude', level=2)

# Explication de l'exemple concret
doc.add_paragraph(
    "Question : Les agriculteurs qui ont accès au crédit agricole sont-ils plus susceptibles de voir leurs revenus augmenter que ceux qui n'y ont pas accès ?\n"
    "Variables :\n"
    "- Variable 1 : Accès au crédit (Oui/Non).\n"
    "- Variable 2 : Amélioration des revenus (Oui/Non).\n"
    "Résultats attendus :\n"
    "- Tableau croisé : Affiche les effectifs observés pour chaque combinaison.\n"
    "- Test du Khi-deux : Affiche la p-value pour déterminer si l'association est significative.\n"
    "Interprétation :\n"
    "- Si la p-value est inférieure à 0.05, vous pouvez conclure que l'accès au crédit est associé à une amélioration des revenus."
)

# Enregistrer le document
doc.save('Guide_Test_Khi_Deux_SPSS.docx')
