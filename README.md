# Symfony UCAOffice365 Bundle

Bundle de connexion d'une application symfony vers l'api développée en interne à l'UCA pour interagir avec les services Microsoft

## Installation

- Lancer la commande `composer require bcedric/uca-office365:dev-main`
- Ajouter le bundle dans `config/bundle.php` :

```
    <?php

    return [
        // ...
        BCedric\UCAOffice365\BCedricUCAOffice365Bundle::class => ['all' => true],
    ];

```

- Pour utiliser le service `BCedric\UCAOffice365\Service\GraphAPI` veuillez définir les variable d'environnement :

```
    GRAPH_TENANT=""
    GRAPH_CLIENT=""
    GRAPH_CLIENT_SECRET=""
```

- Pour utiliser le service `BCedric\UCAOffice365\Service\UCAOffice365` veuillez définir les variable d'environnement :

```
    APIO365_URL=""
    APIO365_LOGIN=""
    APIO365_PASSWORD=""
```

### Service UCAOffice365

| Fonction             | Description                                               |
| -------------------- | --------------------------------------------------------- |
| getUser($uid)        | Retourne les informations concernant l'utilisateur        |
| deleteUser($uid)     | supprime l'utilisateur                                    |
| createUser($uid)     | Ajoute l'utilisateur                                      |
| addBooking($uid)     | Ajoute l'option Booking sur l'utilisateur                 |
| removeBooking($uid)  | Supprime l'option Booking sur l'utilisateur               |
| getCalendarURL($uid) | Renvoie l'URL pour la synchronisation du calendrier Teams |
