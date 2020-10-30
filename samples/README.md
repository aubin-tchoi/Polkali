## Samples

Les scripts suivants sont classés (grossièrement) par difficulté croissante.
Ils n'ont pas réellement pour vocation d'être utilisés en tant que tels (sauf `mailing.gs` et `formating.gs` éventuellement), mais peuvent servir de sample à compléter en fonction du besoin.
Dans ces scripts on trouve souvent en fin de fonction des variables vides, qui sont vouées à être remplies manuellement (il s'agit souvent des IDs des spreadsheets que l'on souhaite faire interagir).

#### import.gs
La fonction `import`permet d'importer des données depuis un autre sheet vers le sheet actif. Il suffit de sélectionner la ligne d'en-tête dans le fichier actif et d'ensuite entrer l'URL et le nom du sheet source lorsque cela est demandé. Les données nouvelles seront ajoutées à la fin du sheet actif. Dans l'idée elle permet de s'affranchir partiellement de la contrainte liée au fait que les projets Apps Script aient besoin d'être liés à des sheets pour pouvoir appeler certaines fonctions.

#### sync.gs
La fonction `sync` permet de synchroniser deux onglets Google Sheets. Couplée à un déclencheur, elle permet de maintenir à jour une copie d'un sheet. Elle nous apprend en outre que les filtres ont un comportement parfois problématique face à la copie d'une range dans sa globalité.

#### update.gs
La fonction `update_onSelec` permet de mettre à jour un fichier destination à partir des informations sélectionnées dans le fichier source. Concrètement une ligne est ajoutée à la fin du fichier destination, et elle est complétée avec les données de la sélection. Vous noterez que l'ordre des informations dans les headers de source et de destination n'a aucune importance, seule les valeurs contenues sont prises en compte (on peut ainsi échanger les colonnes dans un des fichiers, et le script reconstituera l'odre).
Cette fonction trouve de nombreuses applications lorsque l'on modifie différentes choses :
* La manière de choisir les données (ici il s'agit d'une sélection, mais on peut penser à un critère de tri)
* La destination (ici une ligne est ajoutée mais on peut mettre à jour une ligne déjà existante en trouvant une clé primaire)
* Le déclenchement (on peut facilement construire une mise à jour automatique)

#### formating.gs
La fonction `formating` permet de formater les cellules sélectionner selon l'un des formats suivants :
* Majuscule à chaque début de mot puis minuscule (Nom/Prénom)
* Numéro de téléphone sans espaces
* Numéro de téléphone avec espaces
* Date
Cette fonction illustre la très grande facilité que l'on peut avoir à formater de très grands sets de données.

#### archiving.gs
La fonction `archiving` permet de créer une copie du speadsheet courant dans le dossier Drive spécifié.
Cette fonctionnalité peut être utilisée entre autres à des fins d'archivage ou de création de backup.
Elle se montrera proablement assez peu utile en tant que telle, mais vous noterez qu'il est possible de facilement sélectionner les sheets à archiver, de copier seulement une partie des données, de renommer la destination (ce qui peut être utile dans le cadre d'une activité régulière qui fait l'objet de rapports réguliers).

#### stats.gs
Les deux fonctions `stats_merged_fetchAllSheets` et `stats_merged_filtered` permettent de construire des statistiques sur l'ensemble des données présentes dans les différents sheets du spreadsheet courant. Pour chaque type d'information un choix intelligent est fait par la fonction `question_type` pour décider comment sera traitée cette information; sachant qu'elle peut être traitée sous forme de PieChart, de ColumnChart ou présentée dans un tableau récapitulatif. Les diagrammes en question peuvent être ensuite envoyés par mail à une adresse spécifiée, et/ou enregistrés sur le Drive dans un dossier dont on spécifiera l'ID. Ils seront également présentés sur une fenêtre.
/!\ Les headers sont supposés se trouver en première ligne de chaque sheet.

#### mailing.gs
