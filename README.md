# TimecampConverter

Programme qui permet de générer un fichier excel formatté à la sauce Unitee pour le CRA à partir des Reports de Timecamp.

## Prérequis :

1. Cloner le projet 
2. Créer un repertoire de travail où les copies des rapports seront déposés
3. Dans la classe **Program.cs** définir la variable **path** en renseignant le chemin vers votre repertoire de travail.


## Utilisation

1. Ouvrez dans votre navigateur la page suivante : [Timecamp report](https://www.timecamp.com/app#/reports/projects_and_tasks/detailed)
2. Enregistrez la page web dans votre repertoire de travail. Celle-ci devrait avoir le nom suivant : **Reports   TimeCamp.html** 
3. Définissez un nom de fichier de sortie dans la variable **outputFileName** avec une extension xlsx
4. Exécutez le programme afin de générer un fichier Excel à l'image du rapports déposé dans le dossier de travail
5. Il ne reste plus qu'à copier coller le contenu du fichier xlsx dans le fichier CRA Unitee

##Attention

- Pour éviter d'avoir des incohérences, enregistrer votre page html en dézoomant au maximant afin de ne pas avoir de scrollbar vertivale sur le rapport (car les tâches qui ne sont pas visible dans la page ne sont pas contenue dans le code html enregistré)
