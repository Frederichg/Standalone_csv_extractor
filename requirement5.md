Je travaille sur un script Python existant appelé `lmm_analyzer.py`. il est dans le même dossier ici Standalone_csv_extractor//

C'est une interface Tkinter qui permet aux étudiants de charger un fichier Excel, de choisir :

- une colonne d’identifiant animal (Animal ID)
- jusqu’à 3 variables indépendantes (factorielles)
- jusqu’à 6 variables dépendantes

Actuellement, le script fait les choses suivantes pour chaque variable dépendante sélectionnée :

1. Il crée un Q-Q plot et fait un test de Shapiro-Wilk sur **les données brutes** de la VD pour tester la normalité.
2. Il ajuste un LMM gaussien via `statsmodels.formula.api.mixedlm` avec la formule :
   `VD ~ VI1 + VI2 + VI3` et un groupe = AnimalID.
3. Si `statsmodels` n’est pas dispo ou si le LMM échoue, il fait un ANOVA de base.
4. Il exporte les résultats (texte + graphiques) dans un PDF.


Je souhaite mettre à jour (en fait en créer un nouveau appelé glmm_anlayser.py) mon script `lmm_analyzer.py` pour monter en gamme statistique en vue d'une publication scientifique. 

Actuellement, le script utilise des LMM gaussiens classiques pour toutes les variables, ce qui est inapproprié pour mes données de neurosciences.

Voici les modifications structurelles et statistiques à apporter :

1. DÉTECTION AUTOMATIQUE DU MODÈLE :
   Le script doit permettre à l'étudiant d'indiquer si sa VD est un sum ou un time. Souvent c'est déja dans le nom de la  (VD) par son suffixe :

- Si la VD finit par `_sum` : Utiliser un GLMM Negative Binomial (lien log).
- Si la VD finit par `_avg_time` : Utiliser un GLMM Gamma (lien log).

2. IMPLÉMENTATION STATISTIQUE (Famille Gamma pour le temps) :
   Pour les variables `_avg_time`, remplace le `smf.mixedlm` (gaussien) par un modèle capable de gérer une distribution Gamma avec un lien log et des effets aléatoires.

- Note : Dans l'écosystème Python, `statsmodels` permet les GLM Gamma, mais le support des effets aléatoires (GLMM) est limité. Si possible, utilise `statsmodels.genmod.generalized_linear_model.GLM` avec une structure d'équations d'estimation généralisées (GEE) pour l'identifiant animal, ou propose l'intégration de la librairie `pymer4` (qui fait le pont avec R/lme4) si elle est installée.
- L'objectif est d'avoir : `Formula: VD ~ VI1 + VI2 + VI3, Family: Gamma, Link: Log, Random Effect: (1|AnimalID)`.

3. CORRECTION DES DIAGNOSTICS (Normality & QQ-Plots) :
   C'est un point critique. Actuellement, le script teste la normalité des données brutes. C'est une erreur.

- Supprime le test de Shapiro-Wilk sur les données brutes.
- Pour les modèles Gamma et NegBin, la normalité des données n'est pas requise.
- À la place, génère le QQ-Plot sur les **résidus de Pearson** ou les **résidus de déviance** du modèle calculé.
- Si le modèle est un GLMM Gamma, le QQ-plot doit refléter l'adéquation à la distribution choisie, pas à une distribution normale brute.

4. MISE À JOUR DE L'INTERFACE ET DU PDF :

- Modifie la méthode `analyze_variable` pour qu'elle affiche clairement le type de modèle utilisé dans la zone de texte ("Modèle : GLMM Gamma (Log-link)" ou "Modèle : GLMM Negative Binomial").
- Mets à jour l'export PDF pour que les titres des sections reflètent ces modèles.
- Dans la section "Normality Assessment", change le titre par "Model Diagnostics (Residuals Analysis)".

5. ROBUSTESSE :

- Assure-toi que si une valeur est égale à 0 dans un modèle Gamma (ce qui est théoriquement impossible pour cette distribution), le script ajoute un décalage minimal (ex: +0.001) ou gère l'erreur proprement pour ne pas faire planter l'interface des étudiants.

Garde la structure Tkinter actuelle et la logique de boucle sur les variables dépendantes. Ne réécris pas tout le script, concentre-toi sur la logique de `perform_statistical_test` et `create_qq_plot`.
