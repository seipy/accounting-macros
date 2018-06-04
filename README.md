# accounting-macros

Macro that we use to generate contracts, salary sheets, accounting things

# Installation

met ce fichier (le fichier joint) dans ton dossier téléchargements

ouvre le programme terminal

écrit ça : sudo cp ~/Téléchargements/first_try.py /usr/lib/libreoffice/share/Scripts/python/contrats.py

puis entre le mot de passe : amour

# Utilisation :

ouvrir le fichier Dropbox/Thomas/contrat travail/contrat d'engagement/modèle par jérôme.odt

menu outil>macro>gérer les macro> python

macros libreoffice>contrats

dans contrat tu as les 6 fonctions :

info personnelle-> pour rentrer les infos personnelles

toggle_euro -> pour passer l'affichage de euro à CHF et vice versa (note, quand tu passes en euro tu dois définir le taux de change)

salaire_brut_chf -> tu rentres le brut en chf et il fait le reste

salaire_net_chf -> tu rentres le net en chf et il fait le reste

salaire_net_eur -> (tu as compris je pense)

salaire_brut_eur -> (devine...)

Pour utiliser une fonction, tu la sélectionnes en cliquant dessus, et ensuite tu appuies sur le bouton exécuter


# Utilisation plus pratique :

menu affichage>barre d'outil>personnaliser

dans la partie droite tu cliques sur le + vert pour créer une nouvelle barre d'outils

dans la partie à gauche, sous catégorie, tu choisis macro

tu choisis "macros libreoffice>contrats>info_personelle", tu cliques sur la flèche vers la droite pour qu'il passe à droite, tu répètes pour les autres fonctions.

tu cliques OK

tu as une barre d'outils avec les fonctions

Note que quand je l'ai fait ici ça a buggé, il a refusé de mettre toutes les fonctions d'un coup sur une barre d'outils, du coup j'ai mis deux fonctions, j'ai appliqué, j'ai remodifié ma barre d'outil poru en rajouter 2, etc...
