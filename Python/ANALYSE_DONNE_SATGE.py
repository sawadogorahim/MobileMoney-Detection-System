# ==========================================================
# SCRIPT COMPLET - DÉTECTION DE FRAUDE ET PRÉVISION
# ORANGE MONEY - ESPOIR TELECOM
# Auteur : SAWADOGO Rahim
# Filière : TSID - ULBO
# Année : 2026
# ==========================================================

# ----------------------------------------------------------
# IMPORTS
# ----------------------------------------------------------

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
import seaborn as sns
import xgboost
import openpyxl
import warnings
warnings.filterwarnings("ignore")

from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import (accuracy_score,precision_score,recall_score,f1_score,roc_auc_score,confusion_matrix,classification_report,RocCurveDisplay)
from sklearn.preprocessing import StandardScaler
from xgboost import XGBClassifier
from statsmodels.tsa.arima.model import ARIMA
from statsmodels.tsa.stattools import adfuller
from statsmodels.graphics.tsaplots import plot_acf, plot_pacf

# ==========================================================
# PARTIE 1 : CHARGEMENT ET CONTRÔLE DES DONNÉES
# ==========================================================

print("\n" + "="*60)
print("PARTIE 1 : CHARGEMENT ET CONTRÔLE DES DONNÉES")
print("="*60)

df = pd.read_excel("C:/TUTORé/RAPPORT DE STAGE/OrangeMoney_BD_Final.xlsx")

print("\nDimensions du jeu de données :", df.shape)
print("\nAperçu des données :")
print(df.head())
print("\nValeurs manquantes :")
print(df.isnull().sum())

doublons = df.duplicated(subset=["TID"]).sum()
print("\nNombre de doublons (TID) :", doublons)
if doublons > 0:df = df.drop_duplicates(subset=["TID"])
print("Doublons supprimés.")

# Valeur manquante Montant
df["Montant"] = df["Montant"].fillna(df["Montant"].median())
print(df)
print("\nPARTIE 1 TERMINÉE AVEC SUCCÈS")

# ==========================================================
# PARTIE 2 : PRÉTRAITEMENT ET ENCODAGE
# ==========================================================

print("\n" + "="*60)
print("PARTIE 2 : PRÉTRAITEMENT ET ENCODAGE")
print("="*60)

# ----------------------------------------------------------
# Conversion de la date
# ----------------------------------------------------------

df["Date"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")
print(df)
# ----------------------------------------------------------
# Conversion du Solde (texte → numérique)
# ----------------------------------------------------------

df["Solde"] = pd.to_numeric(df["Solde"].astype(str).str.replace(",", "").str.replace(" ", ""),errors="coerce").fillna(0)
print(df)
# ----------------------------------------------------------
# Extraction de l'heure numérique
# ----------------------------------------------------------

df["heure_num"] = df["Heure"].apply(lambda x: int(str(x).split(":")[0]) if pd.notna(x) else 0)
print(df)
# ----------------------------------------------------------
# Tranche horaire CORRIGÉE (00h-5h59 = Nuit)
# ----------------------------------------------------------

def tranche(h):
    if 0 <= h < 6:        # 00h00 à 05h59
        return "Nuit"
    elif 6 <= h < 12:     # 06h00 à 11h59
        return "Matin"
    elif 12 <= h < 18:    # 12h00 à 17h59
        return "Après-midi"
    else:                  # 18h00 à 23h59
        return "Soir"

df["tranche_horaire"] = df["heure_num"].apply(tranche)
print(df)
# ----------------------------------------------------------
# Jour de la semaine réel depuis la Date
# ----------------------------------------------------------

jours = {0: "Lundi", 1: "Mardi", 2: "Mercredi",3: "Jeudi", 4: "Vendredi", 5: "Samedi", 6: "Dimanche"}
df["jour_semaine_reel"] = df["Date"].dt.dayofweek.map(jours)
print(df)
# ----------------------------------------------------------
# Calcul du score de risque et du statut_fraude
# (10 critères pondérés, seuils par percentiles 95/99,
#  seuil de décision = 4 points — cf. Chapitre 2, section 2.3.2)
# ----------------------------------------------------------

seuil_montant     = df["Montant"].quantile(0.95)
seuil_tres_gros   = df["Montant"].quantile(0.99)
seuil_nb_trans    = df["nombre_transactions_jour"].quantile(0.95)
seuil_montant_cum = df["montant_cumule_jour"].quantile(0.95)
seuil_frequence   = df["frequence_client"].quantile(0.95)
print(df)
df["score_risque"] = 0
print(df)
# Critère 1 : Très gros montant (≥ 99e percentile) : +3
df.loc[df["Montant"] >= seuil_tres_gros, "score_risque"] += 3
print(df)
# Critère 2 : Gros montant (95e ≤ Montant < 99e percentile) : +2
df.loc[(df["Montant"] >= seuil_montant) & (df["Montant"] < seuil_tres_gros), "score_risque"] += 2
print(df)
# Critère 3 : Très petit montant (< 500 FCFA) : +1
df.loc[df["Montant"] < 500, "score_risque"] += 1
print(df)
# Critère 4 : Transaction de nuit : +1
df.loc[df["tranche_horaire"] == "Nuit", "score_risque"] += 1
print(df)
# Critère 5 : Nombre élevé de transactions dans la journée (≥ 95e percentile) : +2
df.loc[df["nombre_transactions_jour"] >= seuil_nb_trans, "score_risque"] += 2
print(df)
# Critère 6 : Montant cumulé journalier élevé (≥ 95e percentile) : +2
df.loc[df["montant_cumule_jour"] >= seuil_montant_cum, "score_risque"] += 2
print(df)
# Critère 7 : Client très actif (fréquence ≥ 95e percentile) : +1
df.loc[df["frequence_client"] >= seuil_frequence, "score_risque"] += 1
print(df)
# Critère 8 : Retrait important (Type = "Retrait" et Montant ≥ 95e percentile) : +2
# NB : la casse "Retrait" doit correspondre exactement aux valeurs de la colonne Type
df.loc[(df["Type"] == "Retrait") & (df["Montant"] >= seuil_montant), "score_risque"] += 2
print(df)
# Critère 9 : Transfert important (Type = "Transfert" et Montant ≥ 95e percentile) : +1
df.loc[(df["Type"] == "Transfert") & (df["Montant"] >= seuil_montant), "score_risque"] += 1
print(df)
# Critère 10 : Solde très faible après une transaction importante : +1
df.loc[(df["Solde"] < 100) & (df["Montant"] >= seuil_montant), "score_risque"] += 1
print(df)
# Règle de décision : score >= 4 => fraude
df["statut_fraude"] = (df["score_risque"] >= 4).astype(int)
print(df)
print("\nStatut fraude recalculé :")
print(df["statut_fraude"].value_counts())
print("Taux de fraude :", round(df["statut_fraude"].mean()*100, 2), "%")
print("\nDistribution du score de risque :")
print(df["score_risque"].value_counts().sort_index())

# ----------------------------------------------------------
# Encodage des variables qualitatives
# ----------------------------------------------------------

type_encode    = {"Réception": 0, "Transfert": 1, "Retrait": 2}
tranche_encode = {"Nuit": 0, "Matin": 1, "Après-midi": 2, "Soir": 3}
jour_encode    = {"Lundi": 0, "Mardi": 1, "Mercredi": 2,"Jeudi": 3, "Vendredi": 4, "Samedi": 5, "Dimanche": 6}

df["type_encode"]    = df["Type"].map(type_encode)
df["tranche_encode"] = df["tranche_horaire"].map(tranche_encode)
df["jour_encode"]    = df["jour_semaine_reel"].map(jour_encode)
df["operateur_encode"] = df["Opérateur"].map({"Orange": 0, "Moov Money": 1}).fillna(0)
print(df)
# ----------------------------------------------------------
# Normalisation Min-Max des variables quantitatives
# ----------------------------------------------------------

scaler = StandardScaler()
colonnes_norm = ["Montant", "Solde", "montant_cumule_jour", "frequence_client"]
df[colonnes_norm] = scaler.fit_transform(df[colonnes_norm].fillna(0))
print(df)
# ----------------------------------------------------------
# Sauvegarde
# ----------------------------------------------------------

df.to_excel("OrangeMoney_BD_Prepare.xlsx", index=False)
print("\nFichier préparé sauvegardé : OrangeMoney_BD_Prepare.xlsx")
print("Dimensions :", df.shape)
print("\nPARTIE 2 TERMINÉE AVEC SUCCÈS")

# ==========================================================
# PARTIE 3 : ANALYSE DESCRIPTIVE ET VISUALISATION
# ==========================================================

print("\n" + "="*60)
print("PARTIE 3 : ANALYSE DESCRIPTIVE")
print("="*60)

data = pd.read_excel("OrangeMoney_BD_Prepare.xlsx")
print(data)
print("\nRépartition du statut de fraude :")
print(data["statut_fraude"].value_counts())
print("Taux de fraude :", round(data["statut_fraude"].mean()*100, 2), "%")

print("\nTransactions par type :")
print(data["Type"].value_counts())

print("\nTransactions par tranche horaire :")
print(data["tranche_horaire"].value_counts())

print("\nTransactions par jour de la semaine :")
print(data["jour_semaine_reel"].value_counts())
variables = ["Montant","Solde","nombre_transactions_jour","montant_cumule_jour","frequence_client"]
print(df[variables].describe())
# Figure 1 : Distribution des montants
plt.figure(figsize=(10, 6))
plt.hist(data["Montant"], bins=40, color="steelblue", edgecolor="white")
plt.title("Figure 1 : Distribution des montants", fontsize=14)
plt.xlabel("Montant normalisé")
plt.ylabel("Nombre de transactions")
plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig("Figure1_Distribution_Montants.png", dpi=150)


# Figure 2 : Types de transaction
plt.figure(figsize=(8, 6))
data["Type"].value_counts().plot(kind="bar", color=["#CC0000", "#FF6600", "#FF9900"])
plt.title("Figure 2 : Répartition des types de transaction", fontsize=14)
plt.xlabel("Type")
plt.ylabel("Nombre")
plt.xticks(rotation=0)
plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig("Figure2_Types_Transaction.png", dpi=150)


# Figure 3 : Fraude pie
plt.figure(figsize=(7, 7))
labels = ["Normale (0)", "Frauduleuse (1)"]
sizes  = data["statut_fraude"].value_counts().sort_index()
colors = ["#4CAF50", "#CC0000"]
plt.pie(sizes, labels=labels, autopct="%1.1f%%", colors=colors,
        startangle=90, shadow=True)
plt.title("Figure 3 : Transactions normales et frauduleuses", fontsize=14)
plt.tight_layout()
plt.savefig("Figure3_Fraude_Pie.png", dpi=150)


# Figure 4 : Tranche horaire
plt.figure(figsize=(8, 6))
ordre = ["Nuit", "Matin", "Après-midi", "Soir"]
counts = data["tranche_horaire"].value_counts().reindex(ordre, fill_value=0)
counts.plot(kind="bar", color=["#990000", "#FF9900", "#FF6600", "#CC0000"])
plt.title("Figure 4 : Répartition par tranche horaire", fontsize=14)
plt.xlabel("Tranche horaire")
plt.ylabel("Nombre")
plt.xticks(rotation=0)
plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig("Figure4_Tranches_Horaires.png", dpi=150)


# Figure 5 : Jour de la semaine
plt.figure(figsize=(10, 6))
ordre_jours = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
counts_jours = data["jour_semaine_reel"].value_counts().reindex(ordre_jours, fill_value=0)
counts_jours.plot(kind="bar", color="steelblue")
plt.title("Figure 5 : Répartition par jour de la semaine", fontsize=14)
plt.xlabel("Jour")
plt.ylabel("Nombre")
plt.xticks(rotation=45)
plt.grid(True, alpha=0.3)
plt.tight_layout()

print("Figure 5 sauvegardée.")

# Figure 6 : Boxplot Montant selon fraude
plt.figure(figsize=(8, 6))
plt.boxplot([data[data["statut_fraude"] == 0]["Montant"],data[data["statut_fraude"] == 1]["Montant"]],labels=["Normale", "Frauduleuse"])
plt.title("Figure 6 : Montant selon le statut de fraude", fontsize=14)
plt.ylabel("Montant normalisé")
plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig("Figure6_Boxplot_Fraude.png", dpi=150)


# Figure 7 : Matrice de corrélation
colonnes_corr = ["Montant", "Solde", "nombre_transactions_jour","montant_cumule_jour", "frequence_client","type_encode", "tranche_encode", "jour_encode","statut_fraude"]
plt.figure(figsize=(12, 8))
sns.heatmap(data[colonnes_corr].corr(),annot=True, fmt=".2f",cmap="Blues", linewidths=0.5)
plt.title("Figure 7 : Matrice de corrélation", fontsize=14)
plt.tight_layout()
plt.savefig("Figure7_Matrice_Correlation.png", dpi=150)


print("\nPARTIE 3 TERMINÉE AVEC SUCCÈS")

# ==========================================================
# PARTIE 4 : MACHINE LEARNING
# ==========================================================

print("\n" + "="*60)
print("PARTIE 4 : MACHINE LEARNING")
print("="*60)

X = data[["Montant","Solde","nombre_transactions_jour","montant_cumule_jour","frequence_client","type_encode","tranche_encode","jour_encode","heure_num"]].fillna(0)

y = data["statut_fraude"]

print("\nVariables explicatives :", list(X.columns))
print("Distribution cible :", dict(y.value_counts()))

X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.30, random_state=42, stratify=y)

print("\nApprentissage :", X_train.shape)
print("Test         :", X_test.shape)

# 1. Régression Logistique
print("\n" + "="*30)
print("1. RÉGRESSION LOGISTIQUE")
print("="*30)
lr = LogisticRegression(max_iter=1000, random_state=42)
lr.fit(X_train, y_train)
pred_lr  = lr.predict(X_test)
proba_lr = lr.predict_proba(X_test)[:, 1]
print("Accuracy  :", round(accuracy_score(y_test, pred_lr), 4))
print("Precision :", round(precision_score(y_test, pred_lr, zero_division=0), 4))
print("Recall    :", round(recall_score(y_test, pred_lr, zero_division=0), 4))
print("F1-score  :", round(f1_score(y_test, pred_lr, zero_division=0), 4))
print("ROC-AUC   :", round(roc_auc_score(y_test, proba_lr), 4))

# 2. Random Forest
print("\n" + "="*30)
print("2. RANDOM FOREST")
print("="*30)
rf = RandomForestClassifier(n_estimators=300, random_state=42, class_weight="balanced")
rf.fit(X_train, y_train)
pred_rf  = rf.predict(X_test)
proba_rf = rf.predict_proba(X_test)[:, 1]
print("Accuracy  :", round(accuracy_score(y_test, pred_rf), 4))
print("Precision :", round(precision_score(y_test, pred_rf, zero_division=0), 4))
print("Recall    :", round(recall_score(y_test, pred_rf, zero_division=0), 4))
print("F1-score  :", round(f1_score(y_test, pred_rf, zero_division=0), 4))
print("ROC-AUC   :", round(roc_auc_score(y_test, proba_rf), 4))

# 3. XGBoost
print("\n" + "="*30)
print("3. XGBOOST")
print("="*30)
scale = (y_train == 0).sum() / (y_train == 1).sum()
xgb = XGBClassifier(n_estimators=300, learning_rate=0.1,max_depth=6, random_state=42,eval_metric="logloss", scale_pos_weight=scale)
xgb.fit(X_train, y_train)
pred_xgb  = xgb.predict(X_test)
proba_xgb = xgb.predict_proba(X_test)[:, 1]
print("Accuracy  :", round(accuracy_score(y_test, pred_xgb), 4))
print("Precision :", round(precision_score(y_test, pred_xgb, zero_division=0), 4))
print("Recall    :", round(recall_score(y_test, pred_xgb, zero_division=0), 4))
print("F1-score  :", round(f1_score(y_test, pred_xgb, zero_division=0), 4))
print("ROC-AUC   :", round(roc_auc_score(y_test, proba_xgb), 4))

# Tableau comparatif
resultats = pd.DataFrame({"Modèle": ["Régression Logistique", "Random Forest", "XGBoost"],"Accuracy":  [round(accuracy_score(y_test, p), 4) for p in [pred_lr, pred_rf, pred_xgb]],"Precision": [round(precision_score(y_test, p, zero_division=0), 4) for p in [pred_lr, pred_rf, pred_xgb]],"Recall":    [round(recall_score(y_test, p, zero_division=0), 4) for p in [pred_lr, pred_rf, pred_xgb]],"F1-score":  [round(f1_score(y_test, p, zero_division=0), 4) for p in [pred_lr, pred_rf, pred_xgb]],"ROC-AUC":   [round(roc_auc_score(y_test, p), 4) for p in [proba_lr, proba_rf, proba_xgb]],})
print("\nTableau comparatif :")
print(resultats.to_string(index=False))
resultats.to_excel("Comparaison_Modeles.xlsx", index=False)

# Figure 8 : Comparaison modèles
plt.figure(figsize=(10, 6))
x = np.arange(len(resultats["Modèle"]))
width = 0.15
metrics = ["Accuracy", "Precision", "Recall", "F1-score", "ROC-AUC"]
colors  = ["#CC0000", "#FF6600", "#FF9900", "#4CAF50", "#2196F3"]
for i, (metric, color) in enumerate(zip(metrics, colors)):
 plt.bar(x + i*width, resultats[metric], width, label=metric, color=color)
plt.title("Figure 8 : Comparaison des modèles", fontsize=14)
plt.xticks(x + width*2, resultats["Modèle"], rotation=0)
plt.ylabel("Score")
plt.legend(loc="lower right")
plt.ylim(0, 1.1)
plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig("Figure8_Comparaison_Modeles.png", dpi=150)


# Figure 9 : Matrice de confusion XGBoost
cm = confusion_matrix(y_test, pred_xgb)
plt.figure(figsize=(7, 6))
sns.heatmap(cm, annot=True, fmt="d", cmap="Blues",xticklabels=["Normale", "Frauduleuse"],yticklabels=["Normale", "Frauduleuse"])
plt.title("Figure 9 : Matrice de confusion - XGBoost", fontsize=14)
plt.xlabel("Prédit")
plt.ylabel("Réel")
plt.tight_layout()
plt.savefig("Figure9_Matrice_Confusion.png", dpi=150)


print("\nRapport de classification - XGBoost :")
print(classification_report(y_test, pred_xgb, target_names=["Normale", "Frauduleuse"]))

# Figure 10 : Courbes ROC
plt.figure(figsize=(8, 6))
RocCurveDisplay.from_estimator(lr,  X_test, y_test, name="Régression Logistique", ax=plt.gca())
RocCurveDisplay.from_estimator(rf,  X_test, y_test, name="Random Forest", ax=plt.gca())
RocCurveDisplay.from_estimator(xgb, X_test, y_test, name="XGBoost", ax=plt.gca())
plt.title("Figure 10 : Courbes ROC des trois modèles", fontsize=14)
plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig("Figure10_Courbes_ROC.png", dpi=150)


# Figure 11 : Importance des variables
importance = pd.DataFrame({"Variable": X.columns,"Importance": xgb.feature_importances_}).sort_values(by="Importance", ascending=False)
print("\nImportance des variables (XGBoost) :")
print(importance.to_string(index=False))
plt.figure(figsize=(10, 7))
plt.barh(importance["Variable"], importance["Importance"], color="#CC0000")
plt.title("Figure 11 : Importance des variables - XGBoost", fontsize=14)
plt.xlabel("Importance")
plt.gca().invert_yaxis()
plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig("Figure11_Importance_Variables.png", dpi=150)


print("\nPARTIE 4 TERMINÉE AVEC SUCCÈS")

# ==========================================================
# PARTIE 5 : PRÉVISION AVEC ARIMA
# ==========================================================

print("\n" + "="*60)
print("PARTIE 5 : PRÉVISION DES TRANSACTIONS AVEC ARIMA")
print("="*60)

data = pd.read_excel("OrangeMoney_BD_Prepare.xlsx")
print(data)
data["Date"] = pd.to_datetime(data["Date"])

serie = data.groupby("Date").size()
print(serie)
serie = serie.asfreq("D", fill_value=0)
print(serie)
print("\nSérie temporelle :")
print(serie.head(10))
print("Nombre de jours :", len(serie))

# Figure 12 : Série temporelle
plt.figure(figsize=(14, 6))
plt.plot(serie, color="#CC0000", linewidth=1.5)
plt.title("Figure 12 : Nombre de transactions par jour", fontsize=14)
plt.xlabel("Date")
plt.ylabel("Nombre de transactions")
plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig("Figure12_Serie_Temporelle.png", dpi=150)


# Test ADF
print("\nTest de Dickey-Fuller :")
resultat_adf = adfuller(serie)
print("Statistique ADF :", round(resultat_adf[0], 4))
print("p-value          :", round(resultat_adf[1], 4))
if resultat_adf[1] < 0.05:
    print("La série est STATIONNAIRE.")
    d = 0
else:
    print("La série n'est PAS stationnaire. Différenciation d=1.")
    d = 1

serie_diff = serie.diff().dropna()

# Figure 13 : ACF
plt.figure(figsize=(10, 5))
plot_acf(serie_diff, lags=20)
plt.title("Figure 13 : Fonction d'autocorrélation (ACF)", fontsize=14)
plt.tight_layout()
plt.savefig("Figure13_ACF.png", dpi=150)


# Figure 14 : PACF
plt.figure(figsize=(10, 5))
plot_pacf(serie_diff, lags=20)
plt.title("Figure 14 : Fonction d'autocorrélation partielle (PACF)", fontsize=14)
plt.tight_layout()
plt.savefig("Figure14_PACF.png", dpi=150)


# Sélection automatique paramètres ARIMA
print("\nSélection automatique des paramètres ARIMA...")
meilleur_aic    = np.inf
meilleurs_params = (1, 1, 1)
for p in range(0, 4):
    for q in range(0, 4):
        try:
            res_test = ARIMA(serie, order=(p, d, q)).fit()
            if res_test.aic < meilleur_aic:
                meilleur_aic     = res_test.aic
                meilleurs_params = (p, d, q)
        except:
            continue

print(f"Meilleurs paramètres : ARIMA{meilleurs_params}")
print(f"AIC minimal          : {round(meilleur_aic, 2)}")

# Modèle ARIMA
modele         = ARIMA(serie, order=meilleurs_params)
resultat_arima = modele.fit()
print("\nRésumé ARIMA :")
print(resultat_arima.summary())

# Prévisions 30 jours
prevision = resultat_arima.forecast(steps=30).clip(lower=0)

# Figure 15 : Prévisions
plt.figure(figsize=(14, 6))
plt.plot(serie, label="Historique", color="#CC0000", linewidth=1.5)
plt.plot(prevision, label="Prévisions (30 jours)",color="#FF6600", linewidth=2.5, linestyle="--")
plt.axvline(x=serie.index[-1], color="gray", linestyle=":", linewidth=1)
plt.title("Figure 15 : Prévision des transactions - ARIMA", fontsize=14)
plt.xlabel("Date")
plt.ylabel("Nombre de transactions")
plt.legend()
plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig("Figure15_Previsions_ARIMA.png", dpi=150)


# Sauvegarde prévisions
df_previsions = pd.DataFrame({"Date": pd.date_range(start=serie.index.max() + pd.Timedelta(days=1),periods=30, freq="D"),"Prévision_transactions": prevision.values.round(0).astype(int)})
df_previsions.to_excel("Previsions_ARIMA.xlsx", index=False)
print("\nPrévisions sauvegardées : Previsions_ARIMA.xlsx")
print(df_previsions)

print("\nPARTIE 5 TERMINÉE AVEC SUCCÈS")

# ==========================================================
# RÉCAPITULATIF FINAL
# ==========================================================

print("\n" + "="*60)
print("SCRIPT TERMINÉ AVEC SUCCÈS")
print("="*60)
print("\nFichiers générés :")
print("  - OrangeMoney_BD_Prepare.csv")
print("  - Comparaison_Modeles.xlsx")
print("  - Previsions_ARIMA.xlsx")
print("  - Figure1 à Figure15 (PNG)")
print(f"\nParamètres ARIMA retenus : {meilleurs_params}")
print("="*60)