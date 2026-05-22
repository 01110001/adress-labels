# Générateur d'étiquettes d'adresses

Ce petit projet crée un document Word imprimable (`.docx`) à partir d'un fichier Excel contenant des adresses. Le document produit des blocs découpables pour coller sur des colis.

## Fichier Excel attendu

Le fichier Excel doit contenir une feuille nommée `Table1`.

La première ligne doit contenir les colonnes attendues par le script, sans données personnelles dans le README. Les noms exacts des colonnes sont définis dans `generate_address_labels.py`.

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

## Format produit

- Papier Lettre, orientation portrait.
- 8 blocs d'adresse par page: 2 colonnes x 4 lignes.
- Chaque bloc contient l'adresse du destinataire et une bordure légère pour faciliter le découpage.
