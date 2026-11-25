# Hotmail via Microsoft Graph

Des scripts Python pour consulter une boîte Outlook.com personnelle et nettoyer le dossier Courrier indésirable sous Windows.

## Installation

Dans PowerShell, depuis le dossier du projet, installe les dépendances et crée ta configuration locale.

```powershell
py -m pip install -r requirements.txt
Copy-Item config.exemple.ini config.ini
```

Renseigne l'identifiant d'application Entra et l'adresse de ton compte Microsoft dans `config.ini`. Si le fichier existe déjà, utilise-le directement.

Dans [Microsoft Entra](https://entra.microsoft.com/), inscris une application qui accepte les comptes Microsoft personnels. Dans **Authentification**, ajoute la plateforme **Applications mobiles et de bureau** avec l'URI de redirection `http://localhost`. Dans **Autorisations d'API**, ajoute les autorisations **déléguées** Microsoft Graph nécessaires.

| Script | Autorisations |
| --- | --- |
| `outlook.py` | `Mail.ReadBasic` et `Calendars.Read` |
| `categoriser_reception.py` | `Mail.ReadBasic` |
| `supprimer_indesirables.py` | `Mail.ReadWrite` |

## Utilisation

```powershell
py outlook.py
py categoriser_reception.py
py supprimer_indesirables.py
py supprimer_indesirables.py --auto
```

- `outlook.py` affiche les cinq derniers messages reçus, les cinq derniers indésirables et les rendez-vous des sept prochains jours du calendrier principal.
- Le mode manuel affiche jusqu'à 50 indésirables. Choisis les numéros, puis saisis `OUI` pour les déplacer.
- Le mode `--auto` parcourt tout le dossier Courrier indésirable. Il déplace immédiatement vers **Éléments supprimés** les messages provenant des adresses exactes de `EXPEDITEURS_AUTO`, dans `supprimer_indesirables.py`.

## Tâche planifiée Windows

Exécute d'abord `py supprimer_indesirables.py --connecter` et autorise l'accès dans ton navigateur. Cette commande prépare le cache de connexion chiffré du compte Windows.

Dans le Planificateur de tâches, crée une tâche et choisis la fréquence souhaitée. Configure l'action **Démarrer un programme**.

| Champ | Valeur |
| --- | --- |
| Programme/script | Chemin complet de `python.exe`, obtenu avec `py -c "import sys; print(sys.executable)"` |
| Ajouter des arguments | `"C:\chemin\du\projet\supprimer_indesirables.py" --planifie` |
| Commencer dans | Chemin complet du dossier du projet |

Utilise le même compte Windows que pour `--connecter` et choisis **Exécuter uniquement si l'utilisateur est connecté**. Sélectionne **Ne pas démarrer une nouvelle instance** si la tâche est déjà en cours. Si le journal indique qu'une connexion est requise, relance la commande avec `--connecter`.

## Journal

Le script de nettoyage écrit dans le dossier `logs/` du projet, dans un fichier nommé selon le jour de la semaine (`lundi.log` à `dimanche.log`). Le journal contient les heures d'exécution, les nombres de messages examinés et déplacés et les erreurs. Les exécutions du même jour de la semaine s’ajoutent au même fichier. Chaque fichier conserve deux archives d'environ 250 Ko. Les objets, adresses et jetons restent hors du journal.

```powershell
Get-Content .\logs\lundi.log -Tail 30
```

## Publication sur GitHub

Le dépôt doit inclure les scripts Python, `requirements.txt`, `config.exemple.ini`, `README.md` et `.gitignore`. Le fichier `config.ini` est ignoré. Le dossier `logs/` est ignoré par Git. Le cache de connexion est stocké hors du projet.

## Catégorisation de la boîte de réception

Lance `py categoriser_reception.py` avec l’autorisation déléguée `Mail.ReadBasic`. Le script affiche un aperçu des 50 messages les plus récents, classés selon leur objet et leur expéditeur. Les dix catégories couvrent les finances, les achats, l’administratif, le travail et la formation, les voyages, les notifications, l’entrepreneuriat, l’informatique et l’IA, la santé et les promotions. Les e-mails sans correspondance apparaissent sous À classer. Les règles se modifient dans `REGLES`. Ce mode lit la boîte de réception en lecture seule.
"# Hotmail_via_Microsoft_Graph" 
