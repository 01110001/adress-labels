# Générateur d'étiquettes d'adresses Viabois

Ce petit projet crée un document Word imprimable (`.docx`) à partir d'un fichier Excel contenant des adresses. Le document produit des blocs découpables pour coller sur des colis, avec l'expéditeur Viabois sur chaque bloc.

## Fichier Excel attendu

Le fichier Excel doit contenir une feuille nommée `Table1`.

La première ligne doit contenir exactement ces champs:

| Champ | Description | Exemple |
| --- | --- | --- |
| `street` | Adresse civique / rue du destinataire | `919 rue Moreau` |
| `city` | Ville du destinataire | `Wickham` |
| `state` | Province ou état | `QC` |
| `zip_code` | Code postal | `J0C1S0` |
| `company_name` | Nom de l'entreprise ou du destinataire | `Transport Michel Noel` |

Les colonnes peuvent contenir des accents. Le script imprime les données telles quelles: il ne corrige pas les codes postaux manquants ou mal formatés.

## Installation

Installer les dépendances Python:

```bash
pip install -r requirements.txt
```

## Utilisation

Commande de base:

```bash
python generate_address_labels.py adresses.xlsx etiquettes_adresses_colis_viabois.docx
```

Commande avec une feuille différente:

```bash
python generate_address_labels.py adresses.xlsx etiquettes.docx --sheet Table1
```

## Format produit

- Papier Lettre, orientation portrait.
- 8 blocs d'adresse par page: 2 colonnes x 4 lignes.
- Chaque bloc contient:
  - l'expéditeur Viabois;
  - le destinataire;
  - une bordure légère pour faciliter le découpage.

## Expéditeur inclus

```text
Viabois
311 rue du camionneur
Saint-Isidore, QC G0S 1S0
Canada
```
