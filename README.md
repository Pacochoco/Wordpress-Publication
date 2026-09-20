# SEO Content Pipeline — F31 Blog

Pipeline en deux étapes pour générer des articles SEO pédagogiques sur Excel/Power BI et les publier automatiquement sur le blog WordPress de F31.

## Vue d'ensemble

```
SEO.xlsx (sujets à traiter)
      │
      ▼
generate_articles.py  ──►  génère titre + contenu HTML via GPT-4.1
      │
      ▼
SEO_generated_50_articles_html_corrige.xlsx (articles générés + enrichis manuellement)
      │
      ▼
SEO_publication_wordpress.py  ──►  publie sur WordPress (brouillon ou publié)
```

Le fichier Excel sert de base de données centrale entre les deux scripts : `generate_articles.py` le remplit avec les articles générés, puis (après relecture et enrichissement — liens, images, méta description) `SEO_publication_wordpress.py` s'en sert pour publier sur le site.

## 1. Génération des articles — `generate_articles.py`

Lit une liste de sujets dans un fichier Excel et génère, pour chaque ligne au statut `Ready`, un article structuré via l'API OpenAI (modèle `gpt-4.1`).

**Colonnes Excel attendues en entrée :**
| Colonne | Rôle |
|---|---|
| `Status` | Doit valoir `Ready` pour être traité ; passe à `Done` une fois généré |
| `Nom_sujet` | Sujet de l'article |
| `Expression clé principale` | Mot-clé SEO principal |
| `Template_content` | Structure imposée pour l'article |

**Colonnes générées en sortie :**
- `Article_title` — titre généré
- `Article_corpus` — contenu HTML généré

Le script utilise un appel structuré (`tool_choice` forcé sur une fonction `generate_article`) pour garantir un titre et un contenu HTML bien formés, plutôt qu'un texte libre à parser.

Le style de rédaction est imposé par un exemple complet inclus dans le prompt (structure Introduction / Syntaxe / Paramètres / Exemples / Erreurs fréquentes / Alternatives / Conclusion), assurant une cohérence éditoriale entre tous les articles générés.

## 2. Publication WordPress — `SEO_publication_wordpress.py`

Lit le fichier Excel enrichi et crée les articles correspondants sur WordPress via l'API REST (`wp-json/wp/v2/posts`), avec authentification par mot de passe d'application.

**Fonctionnalités principales :**
- Upload et association d'une **image mise en avant** et d'une **image insérée dans le corps** (positionnable avant le N-ième `<h2>`)
- Remplacement automatique de la première occurrence d'une ancre par un **lien interne ou externe** (jusqu'à 2 liens internes + 1 lien externe)
- Ajout d'une **source externe** en bas d'article si renseignée
- Gestion des **catégories** WordPress (création automatique si inexistante)
- Renseignement des **métadonnées Yoast SEO** (méta description, expression clé)
- **Mode `DRY_RUN`** : permet de simuler une publication sans rien envoyer à WordPress (affiche le payload complet en console)

**Colonnes Excel principales attendues :**
| Colonne | Rôle |
|---|---|
| `Title`, `generated_content` | Titre et contenu de l'article |
| `Status` | Statut de publication WordPress (brouillon / publié) |
| `Category` | Catégorie WordPress |
| `image_featured`, `image_content` | Sources des images (mise en avant / corps) |
| `image_insert_before_heading_number` | Position d'insertion de l'image dans le corps |
| `Lien Interne 1/2 URL` + `ancre`, `Lien Externe 1 URL` + `ancre` | Liens à insérer automatiquement |
| `Expression clé principale`, `Méta description` | Champs Yoast SEO |

> Le script ne traite actuellement que la première ligne du fichier (`df.head(1)`) — utile en phase de test, à ajuster (`df.iterrows()`) pour traiter l'ensemble du fichier en production.

## Installation

```bash
pip install pandas openpyxl openai requests python-dotenv
```

## Configuration requise

Les deux scripts nécessitent des identifiants sensibles qui **ne doivent jamais être codés en dur dans le fichier source**. Copier `.env.example` en `.env` et renseigner :

```bash
# .env (à créer, jamais commité)
OPENAI_API_KEY=...
WORDPRESS_URL=https://f31.com
WORDPRESS_USERNAME=...
WORDPRESS_APP_PASSWORD=...

# Optionnel — sinon valeurs par défaut codées dans les scripts
SEO_EXCEL_PATH=C:\chemin\vers\SEO.xlsx
SEO_GENERATED_EXCEL_PATH=C:\chemin\vers\SEO_generated_50_articles_html_corrige.xlsx
```

Les deux scripts chargent automatiquement ce fichier via `python-dotenv` (`load_dotenv()`), aucune modification supplémentaire n'est nécessaire.

## Utilisation

**Étape 1 — Génération des articles :**
```bash
python generate_articles.py
```
Traite toutes les lignes au statut `Ready` du fichier Excel source, et les marque `Done` une fois l'article généré.

**Étape 2 — Relecture manuelle**
Avant publication, le fichier généré est enrichi manuellement : vérification du contenu, ajout des images, liens internes/externes, méta description.

**Étape 3 — Publication WordPress :**
```bash
python SEO_publication_wordpress.py
```
Vérifier `DRY_RUN = True` pour un premier test sans publication réelle.

## Limites connues / points d'attention

- Le script de publication ne traite que la première ligne du fichier (`df.head(1)`) — limitation volontaire en phase de test
- Aucune gestion de nouvelle tentative (retry) en cas d'échec d'appel API (génération ou publication)
- Le prompt de génération d'article contient un exemple de style unique — pourrait être externalisé dans un fichier séparé pour faciliter les mises à jour éditoriales

## Sécurité

Les identifiants (clé API OpenAI, mot de passe d'application WordPress) doivent être stockés dans un fichier `.env` non commité (voir `.gitignore`), jamais écrits en clair dans le code source.
