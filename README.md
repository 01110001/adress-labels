# Générateur d'étiquettes d'adresses

Ce petit projet crée un document Word imprimable (`.docx`) à partir d'un fichier Excel contenant des adresses. Le document produit des blocs découpables pour coller sur des colis.

## Fichier Excel attendu

Le fichier Excel doit contenir une feuille nommée `Table1`.

La première ligne doit contenir exactement ces champs:

| Champ          | Description                            |
| -------------- | -------------------------------------- |
| `street`       | Adresse civique / rue du destinataire  |
| `city`         | Ville du destinataire                  |
| `state`        | Province ou état                       |
| `zip_code`     | Code postal                            |
| `company_name` | Nom de l'entreprise ou du destinataire |

Les colonnes peuvent contenir des accents. Le script imprime les données telles quelles: il ne corrige pas les codes postaux manquants ou mal formatés.

## Installation

Installer le projet et ses dépendances Python:

```bash
pip install .
```

## Utilisation

Commande de base:

```bash
python generate_address_labels.py adresses.xlsx etiquettes_adresses_colis.docx
```

Commande avec une feuille différente:

```bash
python generate_address_labels.py adresses.xlsx etiquettes.docx --sheet Table1
```

Commande avec un expéditeur à entrer au clavier:

```bash
python generate_address_labels.py adresses.xlsx etiquettes.docx --set-expediteur
```

Avec `--set-expediteur`, le programme demande le nom, l'adresse, la ville, la province ou l'état, le code postal et le pays. Les champs laissés vides sont ignorés.

## Format produit

- Papier Lettre, orientation portrait.
- 8 blocs d'adresse par page: 2 colonnes x 4 lignes.
- Chaque bloc contient l'adresse du destinataire, l'expéditeur seulement si l'option est utilisée, et une bordure légère pour faciliter le découpage.
