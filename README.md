# Excel Invoice Automation

## 1. Présentation du projet

Ce projet Python permet d’automatiser la génération de factures clients à partir de plusieurs fichiers Excel contenant des feuilles de commande.

L’objectif est de remplacer un traitement manuel répétitif par un script capable de :

- Lire plusieurs fichiers Excel (.xlsx)
- Vérifier leur conformité
- Extraire les données des clients
- Regrouper les commandes par client
- Trier les commandes par date
- Générer automatiquement une facture Excel par client

Ce projet s’inscrit dans une logique de traitement de données proche d’un processus ETL (Extraction – Transformation – Chargement).

---

## 2. Objectifs pédagogiques

Ce projet a été réalisé dans le cadre d’un apprentissage personnel afin de :

- Manipuler des fichiers Excel avec Python
- Structurer un projet en plusieurs modules
- Implémenter une logique de validation de données
- Automatiser la production de fichiers
- Utiliser Git et GitHub pour versionner un projet

---

## 3. Technologies utilisées

- Python 3
- openpyxl
- os
- datetime

---

## 4. Structure du projet
\subsection{Structure du projet}

Le projet est organisé selon l’architecture suivante :

\begin{verbatim}
excel-invoice-automation/
|
|-- data/                  (Fichiers Excel d'entrée)
|-- factureclients/        (Factures générées automatiquement)
|
|-- main.py                (Script principal)
|-- traitement_excel.py    (Fonctions de traitement des données)
|-- requirements.txt       (Dépendances du projet)
|-- README.md              (Documentation)
\end{verbatim}

\subsubsection*{Description des composants}

\begin{itemize}

\item \textbf{main.py} : 
Script principal du programme. Il assure :
\begin{itemize}
    \item Le choix du dossier contenant les fichiers Excel
    \item La vérification des fichiers
    \item L'appel des fonctions de traitement
    \item La génération des factures clients
\end{itemize}

\item \textbf{traitement\_excel.py} :
Contient les fonctions de traitement des données :
\begin{itemize}
    \item Vérification de la conformité des fichiers
    \item Extraction des clients
    \item Regroupement des commandes
    \item Structuration des données
\end{itemize}

\item \textbf{data/} :
Dossier contenant les fichiers Excel sources à traiter.

\item \textbf{factureclients/} :
Dossier créé automatiquement pour stocker les factures générées.

\item \textbf{requirements.txt} :
Liste des bibliothèques Python nécessaires à l’exécution du projet.

\end{itemize}
