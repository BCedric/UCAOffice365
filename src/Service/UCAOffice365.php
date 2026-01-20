<?php

namespace BCedric\UCAOffice365\Service;

use Symfony\Component\Cache\Adapter\FilesystemAdapter;
use Symfony\Component\DependencyInjection\Attribute\Autowire;
use Symfony\Contracts\Cache\ItemInterface;
use Symfony\Contracts\HttpClient\HttpClientInterface;

class UCAOffice365
{


    public function __construct(
        private readonly HttpClientInterface $httpClient,
        #[Autowire(env: 'APIO365_URL')] private readonly string $url,
        #[Autowire(env: 'APIO365_LOGIN')] private readonly string $login,
        #[Autowire(env: 'APIO365_PASSWORD')] private readonly string $password,

    ) {}

    public function query($uri, $method = 'GET', $body = [])
    {
        $url = $this->url . $uri;

        $token = $this->getToken();
        $response = $this->httpClient->request($method, $url, [
            'headers' => [
                'Authorization' => 'Bearer ' . $token['access_token']
            ],

        ]);
        return $response->getContent();
    }

    private function getToken()
    {
        $cache = new FilesystemAdapter();
        $cache->clear('my_cache_key');
        return json_decode(
            $cache->get('my_cache_key', function (ItemInterface $item): string {
                $item->expiresAfter(3600);

                $url = $this->url . 'token';
                $response = $this->httpClient->request('POST', $url, [
                    'body' => [
                        'username' => $this->login,
                        'password' => $this->password
                    ]
                ]);
                return $response->getContent();
            }),
            true
        );
    }

    public function getUser($uid)
    {
        $res = json_decode($this->query('user/' . $uid . "/tenant"), true);

        if ($res == null || gettype($res) === 'string') {
            return null;
        }
        return $this->formatUser($res);
    }

    private function formatUser(array $user)
    {
        $user = [...$user, ...$user['TENANT']];
        unset($user['TENANT']);
        return $user;
    }

    public function deleteUser($uid)
    {
        $this->query('user/' . $uid, 'DELETE');
        $user = $this->getUser($uid);
        while ($user == null || $user['status'] != 'deleted') {
            $user = $this->getUser($uid);
            sleep(1);
        }
        return $user;
    }

    public function createUser($uid)
    {

        $this->query('user/' . $uid, 'POST');
        $user = $this->getUser($uid);
        while (
            $user == null || $user['status'] != 'created'
        ) {
            $user = $this->getUser($uid);
            sleep(1);
        }

        return $user;
    }

    public function addBooking(string $uid)
    {
        $this->query('user/' . $uid . "/add_bookings", 'PUT');
        return $this->getUser($uid);
    }

    public function removeBooking(string $uid)
    {
        $this->query('user/' . $uid . "/disable_bookings", 'PUT');
        return $this->getUser($uid);
    }

    public function getCalendarURL(string $uid)
    {
        return json_decode($this->query('user/' . $uid . "/calendar", 'PUT'), true)['data'];
    }
}
