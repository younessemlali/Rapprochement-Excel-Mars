# 🔗 Excel Matcher Beeline

Application Streamlit pour le rapprochement automatique entre fichiers Excel PDF (App 1) et fichiers Excel Beeline.

## 🎯 Fonctionnalités

- ✅ **Upload multiple illimité** pour Excel PDF et Excel Beeline
- ✅ **Rapprochement automatique** par N° commande + Semaine
- ✅ **Validation des montants** avec tolérance configurable
- ✅ **Enrichissement des données** avec noms des intérimaires
- ✅ **Rapport consolidé Excel** avec 5 feuilles d'analyse
- ✅ **Statistiques détaillées** de matching
- ✅ **Interface intuitive** avec visualisations

## 📊 Critères de rapprochement

### **Matching Exact** :
- ✅ N° de commande identique
- ✅ Semaine finissant le identique
- ✅ Montants cohérents (±5% par défaut)

### **Matching Partiel** :
- ✅ N° de commande identique
- ✅ Montants similaires (tolérance élargie)

## 🚀 Installation et utilisation

### 1. Cloner le repository
```bash
git clone https://github.com/votre-username/excel-matcher-beeline.git
cd excel-matcher-beeline
```

### 2. Installer les dépendances
```bash
pip install -r requirements.txt
```

### 3. Lancer l'application
```bash
streamlit run app.py
```

### 4. Utilisation
1. **Uploadez** vos fichiers Excel PDF (issus de l'App 1)
2. **Uploadez** vos fichiers Excel Beeline
3. **Ajustez** la tolérance sur les montants si nécessaire
4. **Lancez** le rapprochement
5. **Vérifiez** les résultats et statistiques
6. **Téléchargez** le rapport consolidé

## 📁 Structure du projet

```
excel-matcher-beeline/
│
├── app.py                 # Application Streamlit principale
├── requirements.txt       # Dépendances Python
├── README.md             # Documentation
└── .streamlit/
    └── config.toml       # Configuration Streamlit
```

## 📊 Format de sortie Excel

Le rapport consolidé contient 5 feuilles :

### 1. Donnees_Consolidees
Toutes les correspondances trouvées avec :
- Données PDF complètes
- Données Beeline complètes (avec noms intérimaires)
- Métadonnées de matching (confiance, écarts)

### 2. Non_Matches_PDF
Données PDF qui n'ont pas trouvé de correspondance.

### 3. Non_Matches_Beeline
Données Beeline qui n'ont pas trouvé de correspondance.

### 4. Statistiques
Métriques détaillées de performance du rapprochement.

### 5. Synthese_Collaborateurs
Analyse par intérimaire avec totaux et écarts.

## 🔧 Paramètres configurables

### Tolérance sur les montants
- **Par défaut** : 5%
- **Plage** : 1% à 20%
- **Usage** : Deux montants sont considérés comme cohérents si leur écart relatif est inférieur à cette tolérance

## 🎯 Algorithme de matching

### Phase 1 : Matching exact
1. Normalisation des N° de commande (suppression caractères non-numériques)
2. Normalisation des semaines (format DD/MM/YYYY)
3. Recherche de correspondances exactes
4. Validation de la cohérence des montants

### Phase 2 : Matching partiel
1. Recherche par N° de commande uniquement
2. Calcul de similarité des montants
3. Sélection du meilleur match (confiance > 40%)

## 📈 Métriques de qualité

- **Taux de match PDF** : % de lignes PDF rapprochées
- **Taux de match Beeline** : % de lignes Beeline rapprochées
- **Confiance moyenne** : Qualité des rapprochements
- **Cohérence des montants** : % de matches avec montants cohérents

## 🛠️ Technologies utilisées

- **Streamlit** : Interface web interactive
- **pandas** : Manipulation des données
- **numpy** : Calculs numériques
- **openpyxl** : Lecture/écriture Excel
- **python-dateutil** : Manipulation des dates

## 🔧 Déploiement sur Streamlit Cloud

### 1. Fork ce repository sur GitHub

### 2. Connecter à Streamlit Cloud
- Aller sur [share.streamlit.io](https://share.streamlit.io)
- Connecter votre compte GitHub
- Déployer l'application

### 3. L'application sera accessible via une URL publique

## 📞 Support

Pour toute question ou problème, créer une issue sur GitHub.

## 📝 License

Ce projet est sous licence MIT.

---

**Excel Matcher Beeline** - Version 1.0 | Rapprochement automatique PDF ↔ Beeline
