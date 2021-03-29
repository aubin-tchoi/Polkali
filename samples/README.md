## Samples

Les scripts suivants sont classés (grossièrement) par difficulté croissante.
Ils n'ont pas réellement pour vocation d'être utilisés en tant que tels.
Dans ces scripts on trouve souvent en fin de fonction des variables vides, qui sont vouées à être remplies manuellement (il s'agit souvent des IDs des spreadsheets que l'on souhaite faire interagir).

#### import.js

La fonction `import`permet d'importer des données depuis un autre sheet vers le sheet actif. Il suffit de sélectionner la ligne d'en-tête dans le fichier actif et d'ensuite entrer l'URL et le nom du sheet source lorsque cela est demandé. Les données nouvelles seront ajoutées à la fin du sheet actif. Dans l'idée elle permet de s'affranchir partiellement de la contrainte liée au fait que les projets Apps Script aient besoin d'être liés à des sheets pour pouvoir appeler certaines fonctions.

#### sync.js

La fonction `sync` permet de synchroniser deux onglets Google Sheets. Couplée à un déclencheur, elle permet de maintenir à jour une copie d'un sheet. Elle nous apprend en outre que les filtres ont un comportement parfois problématique face à la copie d'une range dans sa globalité.

#### update.js

La fonction `updateOnSelec` permet de mettre à jour un fichier destination à partir des informations sélectionnées dans le fichier source. Concrètement une ligne est ajoutée à la fin du fichier destination, et elle est complétée avec les données de la sélection. Vous noterez que l'ordre des informations dans les headers de source et de destination n'a aucune importance, seule les valeurs contenues sont prises en compte (on peut ainsi échanger les colonnes dans un des fichiers, et le script reconstituera l'odre).
Cette fonction trouve de nombreuses applications lorsque l'on modifie différentes choses :

* La manière de choisir les données (ici il s'agit d'une sélection, mais on peut penser à un critère de tri)
* La destination (ici une ligne est ajoutée mais on peut mettre à jour une ligne déjà existante en trouvant une clé primaire)
* Le déclenchement (on peut facilement construire une mise à jour automatique)

#### formating.js

La fonction `formating` permet de formater les cellules sélectionner selon l'un des formats suivants :

* Majuscule à chaque début de mot puis minuscule (Nom/Prénom)
* Numéro de téléphone sans espaces
* Numéro de téléphone avec espaces
* Date

Cette fonction illustre la très grande facilité que l'on peut avoir à formater de très grands sets de données.

#### archiving.js

La fonction `archiving` permet de créer une copie du speadsheet courant dans le dossier Drive spécifié.
Cette fonctionnalité peut être utilisée entre autres à des fins d'archivage ou de création de backup.
Elle se montrera proablement assez peu utile en tant que telle, mais vous noterez qu'il est possible de facilement sélectionner les sheets à archiver, de copier seulement une partie des données, de renommer la destination (ce qui peut être utile dans le cadre d'une activité régulière qui fait l'objet de rapports réguliers).

#### mailToBdd.js

La fonction `addToBdd` permet de lire automatiquement les données du dernier mail reçu (le choix du mail peut être paramétré, et être sujet à une recherche intelligente) pour en extraire des infos à ajouter à une base de données. Le script prévoit le cas où les infos se trouvent dans les pièces jointes, le cas où les infos se trouvent dans le corps du mail est plus simple. En extandant ce script on peut ajuster la manière dont les infos sont ajoutées à la base de données (comme dans d'autres samples, on peut lire son header, convertir les données en un array d'objets et ainsi mettre les infos correspondantes dans la BDD).

#### stats.js

Les deux fonctions `statsMergedFetchAllSheets` et `statsMergedFiltered` permettent de construire des statistiques sur l'ensemble des données présentes dans les différents sheets du spreadsheet courant. Pour chaque type d'information un choix intelligent est fait par la fonction `questionType` pour décider comment sera traitée cette information; sachant qu'elle peut être traitée sous forme de PieChart, de ColumnChart ou présentée dans un tableau récapitulatif. Les diagrammes en question peuvent être ensuite envoyés par mail à une adresse spécifiée, et/ou enregistrés sur le Drive dans un dossier dont on spécifiera l'ID. Ils seront également présentés sur une fenêtre.
/!\ Les headers sont supposés se trouver en première ligne de chaque sheet.

#### genForms.js

La fonction `genForms` permet de générer un Google Forms à partir des données d'un Google Sheet, en créant des sections distinctes pour chaque ligne de données du Sheets. Un trigger `onFormSubmit` est automatiquement ajouté afin de mettre à jour le Google Forms lorsqu'une option est choisie en enlevant du Forms l'option choisie.
Cette fonction considère une disposition particulière des données dans le sheet source, il s'agit de l'adapter à chaque usage en vérifiant chaque getRange.

#### mailing.js

La fonction `mailing` permet d'envoyer des mails personnalisés à un grand nombre de personnes. Cette fonction repose sur différents prérequis :

* On dispose d'un tableau de données contenant les adresses mails à contacter ainsi que les colonnes "Template" et "Date d'envoi du mail" (en lignes les différentes personnes)
* Des modèles Gmail ont été conçus aux alias spécifié dans cette colonne "Template"

La colonne "Template" sert à choisir quel template utiliser pour chaque personne, et la colonne "Date d'envoi du mail" permet de retenir quand le mail a été envoyé, en plus de servir de confirmation (un mail est envoyé uniquement si cette cellule est vide).
Il est possible d'inclure des images dans le modèle (images à ajouter comme dans un mail classique) ainsi que des pièces jointes, ces pièces jointes devant d'abord être upload sur le Drive avant d'être mentionnées dans le modèle par la mention {PJ=`fileURL`}, `fileURL` étant l'URL du fichier à ajouter en pièce jointe (entre guillemets). Cette mention sera supprimée lors de l'envoi.
Il est également possible de personnaliser les mails avec les données du tableau d'entrée, pour ce faire il suffit d'inscrire dans le modèle la mention `{key}` pour qu'elle soit complété pour chaque personne par ce qui se trouve dans la cellule correspondant à la colonne `key` (le header doit contenir la valeur `key`).
N'oubliez pas d'adapter la position du header dans ce script à la configuration du Google Sheets considéré.
