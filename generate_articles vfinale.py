import pandas as pd
import json
import os
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])

file_path = os.environ.get(
    "SEO_EXCEL_PATH",
    r"C:\Users\FrançoisDelepoulle\Pyramid\F31 - Files\Communication\Blog\1. Articles\2. Articles en cours\Automatisation SEO\SEO.xlsx",
)

df = pd.read_excel(file_path)

# Crée les colonnes si elles n'existent pas
if "Article_title" not in df.columns:
    df["Article_title"] = ""

if "Article_corpus" not in df.columns:
    df["Article_corpus"] = ""

for index, row in df.iterrows():

    if str(row["Status"]).strip() != "Ready":
        continue

    topic = row["Nom_sujet"]
    keyword = row["Expression clé principale"]
    template = row["Template_content"]

    system_prompt = """
Tu rédiges des articles SEO pédagogiques sur Excel.
Respecte la structure fournie et copie le style de l'article exemple.
"""

    user_prompt = f"""
Sujet : {topic}
Expression clé : {keyword}

Structure :
{template}

Style à imiter :
# Fonction RECHERCHEV Excel : guide complet

## Introduction
La fonction RECHERCHEV est l’une des fonctions de recherche les plus utilisées dans Excel. Elle permet de trouver une valeur dans la première colonne d’un tableau et de renvoyer une information située sur la même ligne dans une autre colonne.

Pendant de nombreuses années, RECHERCHEV a été la méthode principale pour relier deux tableaux de données dans Excel. On l’utilise notamment pour :
- relier des bases clients
- retrouver un prix produit
- enrichir un export de données
- automatiser des rapports Excel

## Syntaxe de la fonction
RECHERCHEV(valeur_cherchée; table_matrice; no_index_col; [valeur_proche])

## Paramètres expliqués
Présentation claire des paramètres.

## Exemple simple
Exemple basique sur petit tableau.

## Exemple avancé
Cas métier concret.

## Erreurs fréquentes
Erreurs classiques à éviter.

## Alternatives
Autres fonctions proches.

## Conclusion
Résumé clair et pratique.
"""

    response = client.responses.create(
        model="gpt-4.1",
        input=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        tools=[{
            "type": "function",
            "name": "generate_article",
            "description": "Generate a structured SEO article",
            "strict": True,
            "parameters": {
                "type": "object",
                "properties": {
                    "title": {"type": "string"},
                    "content_html": {"type": "string"}
                },
                "required": ["title", "content_html"],
                "additionalProperties": False
            }
        }],
        tool_choice={"type": "function", "name": "generate_article"}
    )

    function_call = next(
        (item for item in response.output if item.type == "function_call"),
        None
    )

    if function_call is None:
        print(f"Ligne {index}: aucun function_call reçu.")
        continue

    result = json.loads(function_call.arguments)

    df.at[index, "Article_title"] = result["title"]
    df.at[index, "Article_corpus"] = result["content_html"]
    df.at[index, "Status"] = "Done"

    print(f"Ligne {index}: article généré pour {topic}")

df.to_excel(file_path, index=False)

print("Articles générés.")