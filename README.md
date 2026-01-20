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

- créer le fichier `config/packages/b_cedric_uca_office365.yaml` avec le contenu suivant :

```
    b_cedric_uca_office365:
        uca_api:
            url: "%env(APIO365_URL)%"
            login: "%env(APIO365_LOGIN)%"
            password: "%env(APIO365_PASSWORD)%"
        graph_api:
            client: "%env(GRAPH_CLIENT)%"
            tenant: "%env(GRAPH_TENANT)%"
            client_secret: "%env(GRAPH_CLIENT_SECRET)%"
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
