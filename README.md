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
- 📊 Génération d’un fichier Excel conforme aux exigences de Sage
- 📁 Traitement en lot de plusieurs factures à la fois
- 🧾 Journalisation des erreurs ou anomalies de lecture

## 🗂️ Exemple de données extraites

- Référence facture
- Code client / site
- Date d’émission
- Montant HT / TVA / TTC
- Période de facturation
- Fournisseur

## 🚀 Installation

Prérequis : Python 3.8+ recommandé. Ensuite :

```bash
git clone https://github.com/ton-utilisateur/smart-turpe.git
cd smart-turpe
pip install -r requirements.txt
