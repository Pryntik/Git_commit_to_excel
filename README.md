# Git_commit_to_excel

## English

### Before use
First, you need to create a commits.txt file (or another custom name) to place in the /assets folder. In this file, place the contents of the following command :
```shell
git log --all --decorate --graph
```

### Description
Converts Git commits to a spreadsheet file. Extracts information (author, date, title, description), groups it by year and month, and then exports it to an XLSX file with appropriate column formatting.

### Presentation
I thought the table format might be easier to proofread. Plus, once in a table format, it can be converted back into other formats, such as CSV, for example.

I didn't search extensively; other than "git pretty format" or a browser extension, I didn't find any options. Being a beginner in Python, I thought it would help me improve my skills in the language (I'm basically a web developer, but I started with Java).

So, at first, I asked ChatGPT to convert my commits to Excel, but the response times were slow (I don't have the pro version, so the number of requests is limited), and large texts weren't supported. So I asked it for a script base, and with Github Copilot on VSCode, I learned a lot. I used it like a teacher, and that's how I learned, corrected, and wrote this Python script. It's not perfect, there are still several possible improvements, thank you for your indulgence and if you have any advice or improvements, don't hesitate !

## Français

### Avant utilisation
Vous devez d'abord créer un fichier commits.txt (ou autre nom personnalisé) a placé dans le dossier /assets. Dans ce fichier, placez le contenu de la commande suivante :
```shell
git log --all --decorate --graph
```

### Description
Convertit les commits Git en fichier tableur. Extrait les informations (auteur, date, titre, description), les regroupe par année et par mois, puis les exporte vers un fichier XLSX avec un formatage de colonnes approprié.

### Présentation
Je me suis dit que le format tableau pourrait être plus simple à relire. En plus, une fois en tableau, il peut être converti à nouveau dans d'autres formats comme en csv par exemple.

Hormis "git pretty format" ou une extension navigateur, je n'ai pas trouvé d'option pour convertir les commits en tableur. Étant débutant en Python, je me suis dit que ça permettrait de m'améliorer sur ce langage (à la base, je suis plutôt développeur web, mais j'ai commencé avec du Java).

Alors au début, je demandais à ChatGPT de convertir mes commits en excel, mais les délais de réponses sont longs (je n'ai pas la version pro donc limitée pour le nombre de requêtes) et puis les gros textes ne sont pas pris en charge. Alors je lui ai demandé une base de script et avec Github copilot sur VSCode j'ai appris plein de choses. Je m'en suis servie comme d'un professeur et c'est comme cela que j'ai appris, corrigé et écrit ce script Python. Il n'est pas parfait, il reste plusieurs améliorations possibles, merci pour votre indulgence et si vous avez des conseils ou des améliorations n'hésitez pas !