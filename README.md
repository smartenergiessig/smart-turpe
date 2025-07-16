# SmartTurpe

**Smart TURPE** est un programme Python conçu pour automatiser la lecture de factures TURPE (Tarif d’Utilisation des Réseaux Publics d'Électricité) au format PDF, en extraire les données clés, et générer automatiquement un fichier Excel ainsi qu'un fichier CSV compatible avec l'importation dans le logiciel de comptabilité **Sage**.

## 📌 Objectif

Faciliter et fiabiliser l'intégration comptable des factures TURPE en :
- évitant la saisie manuelle,
- assurant la cohérence des données extraites,
- et générant un fichier structuré prêt à l'import.

## 🔧 Fonctionnalités principales

- 📥 Lecture automatique de factures TURPE au format PDF
- 🔍 Extraction des informations utiles : numéro de facture, date, montant HT/TTC, code client, période de consommation, etc.
- 📊 Génération d’un fichier Excel et CSV conforme aux exigences de Sage
- 📁 Traitement en lot de plusieurs factures à la fois
- 🧾 Journalisation des erreurs ou anomalies de lecture

## 🗂️ Exemple de données extraites

- Cardi
- Mapping
- Société et/ou
- Etablissement
- Date d'écriture
- Code compte
- N° pièce
- Montant EUR


## 🚀 Installation

Prérequis : Python 3.8+ recommandé. Ensuite :

```bash
git clone https://github.com/ton-utilisateur/smart-turpe.git
cd smart-turpe
pip install -r requirements.txt

## *⚙️ Utilisation*

Il suffit de placer le fichier GestionSPV.xlsx de l'exploitation dans le même fichier que le fichier python ou bien l'exécutable et ensuite d'exécuter le code.
Voir le fichier de Documentation Technique pour la création de l'exécutable (.exe).

