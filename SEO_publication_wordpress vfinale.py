import os
import mimetypes
import requests
import re
import pandas as pd
from requests.auth import HTTPBasicAuth
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

# =========================
# CONFIGURATION
# =========================

WORDPRESS_URL = os.environ["WORDPRESS_URL"]
WORDPRESS_USERNAME = os.environ["WORDPRESS_USERNAME"]
WORDPRESS_APP_PASSWORD = os.environ["WORDPRESS_APP_PASSWORD"]

EXCEL_FILE = os.environ.get(
    "SEO_GENERATED_EXCEL_PATH",
    r"C:\Users\FrancoisDelepoulle\F31\F31 - Documents\Communication\Communication\Blog\1. Articles\2. Articles en cours\Automatisation SEO\SEO_generated_50_articles_html_corrige.xlsx",
)


# =========================
# COLONNES EXCEL UTILISÉES
# =========================

TITLE_COLUMN = "Title"
CONTENT_COLUMN = "generated_content"
STATUS_COLUMN = "Status"
DATE_COLUMN = "date"

CATEGORY_COLUMN = "Category"

# Ancienne colonne éventuelle pour ajouter une source externe en bas d'article
EXTERNAL_LINK_COLUMN = "Lien Externe"

# Images séparées
FEATURED_IMAGE_COLUMN = "image_featured"
CONTENT_IMAGE_COLUMN = "image_content"

FEATURED_IMAGE_ALT_COLUMN = "image_featured_alt"
FEATURED_IMAGE_TITLE_COLUMN = "image_featured_title"

CONTENT_IMAGE_ALT_COLUMN = "image_content_alt"
CONTENT_IMAGE_TITLE_COLUMN = "image_content_title"

IMAGE_INSERT_BEFORE_HEADING_NUMBER_COLUMN = "image_insert_before_heading_number"

# Compatibilité si tu gardes encore ces anciennes colonnes
IMAGE_ALT_COLUMN = "image_alt"
IMAGE_TITLE_COLUMN = "image_title"

KEYPHRASE_COLUMN = "Expression clé principale"
META_DESCRIPTION_COLUMN = "Méta description"

INTERNAL_LINK_1_URL_COLUMN = "Lien Interne 1 URL"
INTERNAL_LINK_1_ANCHOR_COLUMN = "Lien Interne 1 ancre"

INTERNAL_LINK_2_URL_COLUMN = "Lien Interne 2 URL"
INTERNAL_LINK_2_ANCHOR_COLUMN = "Lien Interne 2 ancre"

EXTERNAL_LINK_1_URL_COLUMN = "Lien Externe 1 URL"
EXTERNAL_LINK_1_ANCHOR_COLUMN = "Lien Externe 1 ancre"


# Mode test : True = n’envoie rien à WordPress
DRY_RUN = False

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}


# =========================
# FONCTIONS DE TEST
# =========================

def test_wordpress_posts_endpoint():
    endpoint = f"{WORDPRESS_URL.rstrip('/')}/wp-json/wp/v2/posts"
    print(f"Test endpoint posts authentifié : {endpoint}")

    response = requests.get(
        endpoint,
        auth=HTTPBasicAuth(WORDPRESS_USERNAME, WORDPRESS_APP_PASSWORD),
        headers=HEADERS,
        timeout=30,
    )

    print(f"Code HTTP GET posts : {response.status_code}")


def test_wordpress_auth():
    endpoint = f"{WORDPRESS_URL.rstrip('/')}/wp-json/wp/v2/users/me"
    print(f"Test auth : {endpoint}")

    response = requests.get(
        endpoint,
        auth=HTTPBasicAuth(WORDPRESS_USERNAME, WORDPRESS_APP_PASSWORD),
        headers=HEADERS,
        timeout=30,
    )

    print(f"Code HTTP auth : {response.status_code}")


# =========================
# OUTILS
# =========================

def is_valid_value(value):
    return value is not None and not pd.isna(value) and str(value).strip() != ""


def normalize_text(value):
    if not is_valid_value(value):
        return ""
    return str(value).strip()


def get_row_value(row, df, column_name, default=None):
    if column_name in df.columns and is_valid_value(row.get(column_name)):
        return row.get(column_name)
    return default

def insert_image_in_content(
    content,
    image_url,
    image_alt=None,
    before_heading_number=None,
):
    """
    Insère l'image avant le Nᵉ <h2>.

    - 1 = avant le premier <h2>
    - 2 = avant le deuxième <h2>
    - 3 = avant le troisième <h2>

    Si le numéro est vide, on utilise 1 par défaut.
    Si aucun H2 n'est trouvé, l'image n'est pas insérée.
    """

    if not is_valid_value(image_url):
        print("Image non insérée : image_url vide.")
        return content

    if not is_valid_value(before_heading_number):
        print("Numéro H2 vide : utilisation de 1 par défaut.")
        heading_number = 1
    else:
        try:
            heading_number = int(float(before_heading_number))
        except ValueError:
            print(f"Numéro H2 invalide : {before_heading_number}. Utilisation de 1 par défaut.")
            heading_number = 1

    if heading_number <= 0:
        print(f"Numéro H2 invalide : {heading_number}. Utilisation de 1 par défaut.")
        heading_number = 1

    image_alt = normalize_text(image_alt)

    image_html = f"""
<!-- wp:image {{"sizeSlug":"large"}} -->
<figure class="wp-block-image size-large">
  <img src="{image_url}" alt="{image_alt}" />
</figure>
<!-- /wp:image -->
"""

    content_lower = content.lower()

    positions = []
    search_start = 0

    while True:
        position = content_lower.find("<h2", search_start)

        if position == -1:
            break

        positions.append(position)
        search_start = position + 3

    print(f"H2 détectés dans le contenu : {len(positions)}")
    print(f"Insertion demandée avant H2 n° : {heading_number}")

    if len(positions) == 0:
        print("Image non insérée : aucun <h2> trouvé dans le contenu.")
        print("Début du contenu reçu par Python :")
        print(repr(content[:500]))
        return content

    if len(positions) < heading_number:
        print(f"Image non insérée : H2 n°{heading_number} introuvable.")
        print(f"Nombre de H2 disponibles : {len(positions)}")
        return content

    insert_position = positions[heading_number - 1]

    return (
        content[:insert_position]
        + image_html
        + "\n"
        + content[insert_position:]
    )

def replace_first_occurrence_with_link(content, anchor, url, external=False):
    """
    Remplace la première occurrence d'une ancre par un lien HTML.

    Exemple interne :
    Power BI -> <a href="https://f31.com/...">Power BI</a>

    Exemple externe :
    Microsoft 365 -> <a href="https://support.microsoft.com/..." target="_blank" rel="noopener noreferrer">Microsoft 365</a>
    """

    if not is_valid_value(anchor) or not is_valid_value(url):
        return content

    anchor = normalize_text(anchor)
    url = normalize_text(url)

    if anchor not in content:
        print(f"Ancre introuvable dans le contenu : {anchor}")
        return content

    if external:
        link_html = f'<a href="{url}" target="_blank" rel="noopener noreferrer">{anchor}</a>'
    else:
        link_html = f'<a href="{url}">{anchor}</a>'

    return content.replace(anchor, link_html, 1)


def append_external_link_to_content(content, external_link=None):
    """
    Ajoute une source externe en bas d'article.
    À utiliser seulement si tu veux aussi afficher une source en fin d'article.
    """

    if not is_valid_value(external_link):
        return content

    external_link = normalize_text(external_link)

    external_html = f"""
<hr>
<p>
  Source externe :
  <a href="{external_link}" target="_blank" rel="noopener noreferrer">
    {external_link}
  </a>
</p>
"""

    return f"{content}\n\n{external_html}"


def get_status_from_row(row):
    if STATUS_COLUMN in row and is_valid_value(row.get(STATUS_COLUMN)):
        status = normalize_text(row.get(STATUS_COLUMN)).lower()

        status_mapping = {
            "ready": "draft",
            "draft": "draft",
            "publish": "publish",
            "published": "publish",
            "pending": "pending",
            "private": "private",
            "future": "future",
        }

        return status_mapping.get(status, "draft")

    return "draft"


# =========================
# CATÉGORIES WORDPRESS
# =========================

def get_or_create_category_id(category_name):
    category_name = normalize_text(category_name)

    if not category_name:
        return None

    endpoint = f"{WORDPRESS_URL.rstrip('/')}/wp-json/wp/v2/categories"

    search_response = requests.get(
        endpoint,
        params={"search": category_name},
        auth=HTTPBasicAuth(WORDPRESS_USERNAME, WORDPRESS_APP_PASSWORD),
        headers=HEADERS,
        timeout=30,
    )

    if search_response.status_code != 200:
        print(f"Erreur recherche catégorie : {category_name}")
        print(f"Code HTTP : {search_response.status_code}")
        print(search_response.text)
        return None

    categories = search_response.json()

    for category in categories:
        if category.get("name", "").strip().lower() == category_name.lower():
            return category.get("id")

    if DRY_RUN:
        print(f"[DRY RUN] Catégorie à créer : {category_name}")
        return None

    create_response = requests.post(
        endpoint,
        json={"name": category_name},
        auth=HTTPBasicAuth(WORDPRESS_USERNAME, WORDPRESS_APP_PASSWORD),
        headers=HEADERS,
        timeout=30,
    )

    if create_response.status_code not in [200, 201]:
        print(f"Erreur création catégorie : {category_name}")
        print(f"Code HTTP : {create_response.status_code}")
        print(create_response.text)
        return None

    category_data = create_response.json()
    print(f"Catégorie créée : {category_name} — ID {category_data.get('id')}")

    return category_data.get("id")


# =========================
# MÉDIAS WORDPRESS
# =========================

def upload_image_to_wordpress(image_value, image_alt=None, image_title=None):
    image_value = normalize_text(image_value)

    if not image_value:
        return None

    if DRY_RUN:
        print(f"[DRY RUN] Image à uploader : {image_value}")
        return None

    if image_value.startswith("http://") or image_value.startswith("https://"):
        local_path = download_image_temporarily(image_value)
    else:
        local_path = image_value

    if not os.path.exists(local_path):
        print(f"Image introuvable : {local_path}")
        return None

    filename = os.path.basename(local_path)
    mime_type, _ = mimetypes.guess_type(local_path)

    if mime_type is None:
        mime_type = "image/jpeg"

    media_endpoint = f"{WORDPRESS_URL.rstrip('/')}/wp-json/wp/v2/media"

    with open(local_path, "rb") as image_file:
        media_headers = {
            **HEADERS,
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Content-Type": mime_type,
        }

        response = requests.post(
            media_endpoint,
            data=image_file,
            auth=HTTPBasicAuth(WORDPRESS_USERNAME, WORDPRESS_APP_PASSWORD),
            headers=media_headers,
            timeout=60,
        )

    if response.status_code not in [200, 201]:
        print(f"Erreur upload image : {image_value}")
        print(f"Code HTTP : {response.status_code}")
        print(response.text)
        return None

    media_data = response.json()
    media_id = media_data.get("id")
    media_url = media_data.get("source_url")

    print(f"Image uploadée : {filename} — ID {media_id}")

    update_media_metadata(
        media_id=media_id,
        alt_text=normalize_text(image_alt),
        title=normalize_text(image_title),
    )

    return {
        "id": media_id,
        "url": media_url,
    }


def download_image_temporarily(image_url):
    temp_dir = "temp_images"
    os.makedirs(temp_dir, exist_ok=True)

    filename = image_url.split("/")[-1].split("?")[0]

    if not filename:
        filename = "image.jpg"

    local_path = os.path.join(temp_dir, filename)

    response = requests.get(
        image_url,
        headers=HEADERS,
        timeout=60,
    )

    if response.status_code != 200:
        raise ValueError(f"Impossible de télécharger l'image : {image_url}")

    with open(local_path, "wb") as file:
        file.write(response.content)

    return local_path


def update_media_metadata(media_id, alt_text=None, title=None):
    if not media_id:
        return

    payload = {}

    if is_valid_value(alt_text):
        payload["alt_text"] = alt_text

    if is_valid_value(title):
        payload["title"] = title

    if not payload:
        return

    endpoint = f"{WORDPRESS_URL.rstrip('/')}/wp-json/wp/v2/media/{media_id}"

    response = requests.post(
        endpoint,
        json=payload,
        auth=HTTPBasicAuth(WORDPRESS_USERNAME, WORDPRESS_APP_PASSWORD),
        headers=HEADERS,
        timeout=30,
    )

    if response.status_code not in [200, 201]:
        print(f"Erreur mise à jour métadonnées image ID {media_id}")
        print(f"Code HTTP : {response.status_code}")
        print(response.text)
        return

    print(f"Métadonnées image mises à jour — ID {media_id}")


# =========================
# CRÉATION ARTICLE
# =========================

def create_wordpress_draft(
    title,
    content,
    status="draft",
    category_id=None,
    featured_media_id=None,
    meta_description=None,
    focus_keyphrase=None,
):
    endpoint = f"{WORDPRESS_URL.rstrip('/')}/wp-json/wp/v2/posts"

    payload = {
        "title": title,
        "content": content,
        "status": status,
    }

    if category_id:
        payload["categories"] = [category_id]

    if featured_media_id:
        payload["featured_media"] = featured_media_id

    yoast_meta = {}

    if is_valid_value(meta_description):
        yoast_meta["_yoast_wpseo_metadesc"] = normalize_text(meta_description)

    if is_valid_value(focus_keyphrase):
        yoast_meta["_yoast_wpseo_focuskw"] = normalize_text(focus_keyphrase)

    if yoast_meta:
        payload["meta"] = yoast_meta

    if DRY_RUN:
        print("\n[DRY RUN] Article prêt à envoyer :")
        print(f"Titre : {title}")
        print(f"Status : {status}")
        print(f"Catégorie ID : {category_id}")
        print(f"Image mise en avant ID : {featured_media_id}")
        print(f"Yoast meta : {yoast_meta}")
        print(f"Contenu aperçu : {content[:500]}")
        return None

    response = requests.post(
        endpoint,
        json=payload,
        auth=HTTPBasicAuth(WORDPRESS_USERNAME, WORDPRESS_APP_PASSWORD),
        headers=HEADERS,
        timeout=30,
    )

    if response.status_code not in [200, 201]:
        print(f"\nErreur pour l’article : {title}")
        print(f"Code HTTP : {response.status_code}")
        print(response.text)
        return None

    data = response.json()

    print(f"\nArticle créé : {title}")
    print(f"Status : {status}")
    print(f"ID WordPress : {data.get('id')}")
    print(f"URL : {data.get('link')}")

    return data


# =========================
# EXÉCUTION
# =========================

def main():
    # test_wordpress_posts_endpoint()
    # test_wordpress_auth()

    df = pd.read_excel(EXCEL_FILE)

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.replace("\u00a0", " ", regex=False)
        .str.replace("\n", " ", regex=False)
        .str.replace("\r", " ", regex=False)
    )

    print("Colonnes détectées dans Excel :")
    for col in df.columns:
        print(f"- {repr(col)}")

    required_columns = [TITLE_COLUMN, CONTENT_COLUMN]

    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Colonne manquante dans l’Excel : {col}")

    for index, row in df.head(1).iterrows():
        title = row.get(TITLE_COLUMN)
        content = row.get(CONTENT_COLUMN)

        if pd.isna(title) or pd.isna(content):
            print(f"Ligne {index + 2} ignorée : titre ou contenu manquant")
            continue

        title = normalize_text(title)
        content = normalize_text(content)

        status = get_status_from_row(row)

        category_id = None
        if CATEGORY_COLUMN in df.columns and is_valid_value(row.get(CATEGORY_COLUMN)):
            category_id = get_or_create_category_id(row.get(CATEGORY_COLUMN))

        # =========================
        # IMAGE MISE EN AVANT
        # =========================

        featured_media_id = None

        featured_image_alt = get_row_value(
            row,
            df,
            FEATURED_IMAGE_ALT_COLUMN,
            default=get_row_value(row, df, IMAGE_ALT_COLUMN, default=title),
        )

        featured_image_title = get_row_value(
            row,
            df,
            FEATURED_IMAGE_TITLE_COLUMN,
            default=get_row_value(row, df, IMAGE_TITLE_COLUMN, default=title),
        )

        if FEATURED_IMAGE_COLUMN in df.columns and is_valid_value(row.get(FEATURED_IMAGE_COLUMN)):
            uploaded_featured_image = upload_image_to_wordpress(
                image_value=row.get(FEATURED_IMAGE_COLUMN),
                image_alt=featured_image_alt,
                image_title=featured_image_title,
            )

            if uploaded_featured_image:
                featured_media_id = uploaded_featured_image.get("id")

        # =========================
        # IMAGE DANS LE CORPS
        # =========================

        content_image_url = None

        content_image_alt = get_row_value(
            row,
            df,
            CONTENT_IMAGE_ALT_COLUMN,
            default=get_row_value(row, df, IMAGE_ALT_COLUMN, default=title),
        )

        content_image_title = get_row_value(
            row,
            df,
            CONTENT_IMAGE_TITLE_COLUMN,
            default=get_row_value(row, df, IMAGE_TITLE_COLUMN, default=title),
        )

        if CONTENT_IMAGE_COLUMN in df.columns and is_valid_value(row.get(CONTENT_IMAGE_COLUMN)):
            uploaded_content_image = upload_image_to_wordpress(
                image_value=row.get(CONTENT_IMAGE_COLUMN),
                image_alt=content_image_alt,
                image_title=content_image_title,
            )

            if uploaded_content_image:
                content_image_url = uploaded_content_image.get("url")

        image_insert_before_heading_number = (
            row.get(IMAGE_INSERT_BEFORE_HEADING_NUMBER_COLUMN)
            if IMAGE_INSERT_BEFORE_HEADING_NUMBER_COLUMN in df.columns
            else None
        )

        # =========================
        # LIENS
        # =========================

        external_link = row.get(EXTERNAL_LINK_COLUMN) if EXTERNAL_LINK_COLUMN in df.columns else None

        internal_link_1_anchor = row.get(INTERNAL_LINK_1_ANCHOR_COLUMN) if INTERNAL_LINK_1_ANCHOR_COLUMN in df.columns else None
        internal_link_1_url = row.get(INTERNAL_LINK_1_URL_COLUMN) if INTERNAL_LINK_1_URL_COLUMN in df.columns else None

        internal_link_2_anchor = row.get(INTERNAL_LINK_2_ANCHOR_COLUMN) if INTERNAL_LINK_2_ANCHOR_COLUMN in df.columns else None
        internal_link_2_url = row.get(INTERNAL_LINK_2_URL_COLUMN) if INTERNAL_LINK_2_URL_COLUMN in df.columns else None

        external_link_1_anchor = row.get(EXTERNAL_LINK_1_ANCHOR_COLUMN) if EXTERNAL_LINK_1_ANCHOR_COLUMN in df.columns else None
        external_link_1_url = row.get(EXTERNAL_LINK_1_URL_COLUMN) if EXTERNAL_LINK_1_URL_COLUMN in df.columns else None

        meta_description = row.get(META_DESCRIPTION_COLUMN) if META_DESCRIPTION_COLUMN in df.columns else None
        focus_keyphrase = row.get(KEYPHRASE_COLUMN) if KEYPHRASE_COLUMN in df.columns else None

        final_content = content

        final_content = replace_first_occurrence_with_link(
            final_content,
            internal_link_1_anchor,
            internal_link_1_url,
            external=False,
        )

        final_content = replace_first_occurrence_with_link(
            final_content,
            internal_link_2_anchor,
            internal_link_2_url,
            external=False,
        )

        final_content = replace_first_occurrence_with_link(
            final_content,
            external_link_1_anchor,
            external_link_1_url,
            external=True,
        )

        final_content = insert_image_in_content(
            content=final_content,
            image_url=content_image_url,
            image_alt=content_image_alt,
            before_heading_number=image_insert_before_heading_number,
        )

        final_content = append_external_link_to_content(
            content=final_content,
            external_link=external_link,
        )
        print("DEBUG image_insert_before_heading_number :", repr(image_insert_before_heading_number))
        print("DEBUG content_image_url :", repr(content_image_url))
        print("DEBUG début final_content :", repr(final_content[:300]))
        print("\n--- DEBUG LIGNE EXCEL ---")
        print(f"Titre : {title}")
        print(f"Status : {status}")
        print(f"Catégorie brute : {row.get(CATEGORY_COLUMN) if CATEGORY_COLUMN in df.columns else 'COLONNE ABSENTE'}")
        print(f"Catégorie ID : {category_id}")
        print(f"Image mise en avant brute : {row.get(FEATURED_IMAGE_COLUMN) if FEATURED_IMAGE_COLUMN in df.columns else 'COLONNE ABSENTE'}")
        print(f"Image mise en avant ID : {featured_media_id}")
        print(f"Image contenu brute : {row.get(CONTENT_IMAGE_COLUMN) if CONTENT_IMAGE_COLUMN in df.columns else 'COLONNE ABSENTE'}")
        print(f"Image contenu URL : {content_image_url if is_valid_value(content_image_url) else 'VIDE OU ABSENT'}")
        print(f"Insertion image avant H2 n° : {image_insert_before_heading_number if is_valid_value(image_insert_before_heading_number) else 'NON RENSEIGNÉ'}")
        print(f"Lien externe bas d'article : {external_link if is_valid_value(external_link) else 'VIDE OU ABSENT'}")
        print(f"Lien interne 1 : {internal_link_1_anchor} -> {internal_link_1_url}")
        print(f"Lien interne 2 : {internal_link_2_anchor} -> {internal_link_2_url}")
        print(f"Lien externe 1 : {external_link_1_anchor} -> {external_link_1_url}")
        print(f"Expression clé : {focus_keyphrase if is_valid_value(focus_keyphrase) else 'VIDE OU ABSENT'}")
        print(f"Méta description : {meta_description if is_valid_value(meta_description) else 'VIDE OU ABSENT'}")

        create_wordpress_draft(
            title=title,
            content=final_content,
            status=status,
            category_id=category_id,
            featured_media_id=featured_media_id,
            meta_description=meta_description,
            focus_keyphrase=focus_keyphrase,
        )


if __name__ == "__main__":
    main()