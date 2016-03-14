# Universal ROM Cleaner 

Ce logiciel vous permet de nettoyer vos romsets facilement en se basant sur les attributs entre parenthèses et entre crochets de vos fichiers.

- Il est multilingue (Anglais/Français)
- Il se veut simple d’utilisation (tout fonctionne à la souris par de simple glissez déplacez)

__Comment ça marche :__ 

1 - Décompressez l’archive et lancez le fichier EXE.  
2 - Dans le menu Fichier/Chargez les roms, choisissez le répertoire de votre set à nettoyer.  
3 - Dans le tableau de gauche, ordonnées vos préférences (ce que vous souhaitez garder le plus en haut, le moins en base)  
4 - Déplacez dans le tableau en haut a droite ce que vous ne souhaitez pas conserver  
5 - Déplacez dans le tableau en bas à droite ce qui ne doit pas être pris en compte dans le tris  
6 - Dans le menu action vous pouvez soit simuler le tri (vous aurez alors un fichier texte avec OK ce qui est gardé, KO ce qui n’est pas gardé) ou directement faire le nettoyage  
7 - C’est finit, vous avez un beau romset tout propre pour votre recalbox)  

__Informations supplémentaire :__
Le soft ne supprime rien, il ne fait que déplacer dans un répertoire “ROM_CLEAN” les roms sélectionnées et triées

![Universal ROM Cleaner Interface](http://zupimages.net/up/16/10/xd29.jpg "Interface")

Dans l’exemple de la copie d’écran voici ce qu’il fait si vous avez :

- mario (USA) (En,Fr,De).zip
- mario (Europe) (En,Fr,De).zip
- mario (Europe) (En,Fr,De,Es).zip
- mario (Europe) (En,Fr,De) (Beta).zip
- mario (Japan).zip  
Le seul fichier gardé sera “mario (Europe) (En,Fr,De).zip”

Parce que : je lui demande de ne pas garder les “Beta” ou les “Japan” et que dans l’ordre, je garde les fichier “Europe” avant les fichier “USA” et que après ce premier choix, je garde en priorité ceux en “En,Fr,De” avant ceux en “En,Fr,De,Es”

N'hesitez pas à faire des simulation si vous n’êtes pas sur de votre choix 😉
