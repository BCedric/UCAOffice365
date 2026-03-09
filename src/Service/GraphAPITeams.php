<?php

namespace BCedric\UCAOffice365\Service;

use Exception;
use Microsoft\Graph\Core\Authentication\GraphPhpLeagueAccessTokenProvider;
use Microsoft\Graph\Core\GraphClientFactory;
use Microsoft\Graph\Generated\Models\OnlineMeeting;
use Microsoft\Graph\GraphRequestAdapter;
use Microsoft\Graph\GraphServiceClient;
use Microsoft\Kiota\Authentication\Oauth\ClientCredentialContext;
use Microsoft\Kiota\Authentication\PhpLeagueAuthenticationProvider;
use Symfony\Component\DependencyInjection\Attribute\Autowire;
use GuzzleHttp\Client as GuzzleClient;
use League\OAuth2\Client\Provider\GenericProvider;

class GraphAPITeams
{
    private ?GraphServiceClient $graphServiceClient = null;
    private ?\Microsoft\Graph\Graph $legacyGraph = null;
    private ?string $legacyAccessToken = null;
    private int $legacyAccessTokenExpireAt = 0;
    public string $userUrlPrefix = 'https://graph.microsoft.com/v1.0/users/';

    public function __construct(
        #[Autowire(env: 'GRAPH_TENANT')] private readonly string $tenantId,
        #[Autowire(env: 'GRAPH_CLIENT')] private readonly string $clientId,
        #[Autowire(env: 'GRAPH_CLIENT_SECRET')] private readonly string $clientSecret,
        #[Autowire(env: 'PROXY_URL')] private readonly ?string $proxyUrl = null,
    ) {
    }

    public function getGraphServiceClient(): GraphServiceClient
    {
        if (is_null($this->graphServiceClient)) {
            $tokenRequestContext = new ClientCredentialContext(
                $this->tenantId,
                $this->clientId,
                $this->clientSecret
            );

            $config = [
                'verify' => true,
                'timeout' => 30,
            ];

            if (!empty($this->proxyUrl)) {
                $config['proxy'] = [
                    'http' => $this->proxyUrl,
                    'https' => $this->proxyUrl,
                ];
            }

            $httpClient = GraphClientFactory::createWithConfig($config);
            $guzzleClientForOAuth = new GuzzleClient($config);
            $providerOptions = [
                'clientId' => $this->clientId,
                'clientSecret' => $this->clientSecret,
                'redirectUri' => '',
                'urlAuthorize' => "https://login.microsoftonline.com/{$this->tenantId}/oauth2/v2.0/authorize",
                'urlAccessToken' => "https://login.microsoftonline.com/{$this->tenantId}/oauth2/v2.0/token",
                'urlResourceOwnerDetails' => 'https://graph.microsoft.com/oidc/userinfo',
                'scopes' => 'https://graph.microsoft.com/.default',
                'verify' => true,
                'timeout' => 30,
            ];

            if (!empty($this->proxyUrl)) {
                $providerOptions['proxy'] = [
                    'http' => $this->proxyUrl,
                    'https' => $this->proxyUrl,
                ];
            }

            $oauthProvider = new GenericProvider($providerOptions, [
                'httpClient' => $guzzleClientForOAuth,
            ]);

            $accessTokenProvider = new GraphPhpLeagueAccessTokenProvider(
                $tokenRequestContext,
                ['https://graph.microsoft.com/.default'],
                \Microsoft\Graph\Core\NationalCloud::GLOBAL,
                null,
                $oauthProvider
            );

            $authProvider = PhpLeagueAuthenticationProvider::createWithAccessTokenProvider($accessTokenProvider);
            $this->graphServiceClient = GraphServiceClient::createWithRequestAdapter(
                new GraphRequestAdapter(
                    $authProvider,
                    $httpClient
                )
            );
        }

        return $this->graphServiceClient;
    }

    public function getLegacyGraph(string $version = 'v1.0')
    {
        if ($this->legacyGraph === null) {
            $this->legacyGraph = new \Microsoft\Graph\Graph();
            $this->legacyGraph->setAccessToken($this->getLegacyAccessToken());
        } else if ($this->legacyAccessTokenExpireAt <= (time() + 30)) {
            $this->legacyGraph->setAccessToken($this->getLegacyAccessToken());
        }

        $this->legacyGraph->setApiVersion($version);
        return $this->legacyGraph;
    }

    public function legacyRequest(
        string $method,
        string $url,
        array $body = [],
        string $version = 'v1.0',
        ?string $returntype = null
    ): mixed {
        $graph = $this->getLegacyGraph($version);
        $path = $this->normalizeGraphPath($url);

        $request = $graph->createRequest(strtoupper($method), $path);
        if (!empty($returntype)) {
            $request->setReturnType($returntype);
        }
        if (!empty($body)) {
            $request->attachBody($body);
        }

        return $request->execute();
    }

    public function addGroupMemberByUserId(string $groupId, string $userId): mixed
    {
        $data = [
            '@odata.id' => $this->userUrlPrefix . $userId,
        ];

        try {
            return $this->legacyRequest('POST', "groups/$groupId/members/\$ref", $data);
        } catch (\Throwable $e) {
            throw new Exception($e->getMessage(), (int)$e->getCode(), $e);
        }
    }

    public function listGroupMemberIdSet(string $groupId): array
    {
        $client = $this->getGraphServiceClient();
        try {
            $set = [];
            $members = $client->groups()->byGroupId($groupId)->members()->get()->wait();
            if (!empty($members) && method_exists($members, 'getValue')) {
                foreach (($members->getValue() ?? []) as $member) {
                    $id = method_exists($member, 'getId') ? $member->getId() : null;
                    if (!empty($id)) {
                        $set[(string)$id] = true;
                    }
                }
            }

            return $set;
        } catch (\Throwable $e) {
            // Keep a robust fallback for older/generated SDK edge cases.
            $set = [];
            $graph = $this->getLegacyGraph('v1.0');
            $path = "/groups/$groupId/members?\$select=id&\$top=999";

            while (!empty($path)) {
                $response = $graph->createRequest('GET', $path)->execute();
                $data = $this->graphResponseToArray($response);

                foreach (($data['value'] ?? []) as $member) {
                    $id = $member['id'] ?? null;
                    if (!empty($id)) {
                        $set[(string)$id] = true;
                    }
                }

                $path = $data['@odata.nextLink'] ?? null;
            }

            return $set;
        }
    }

    public function createOnlineMeetingExtended(
        string $userId,
        string $subject,
        ?int $startTimestamp = null,
        ?int $endTimestamp = null,
        bool $recordAutomatically = false,
        array $coOrganizerIds = [],
        ?string $roomEmail = null
    ): mixed {
        $body = [
            'subject' => $subject,
        ];

        if (!empty($startTimestamp)) {
            $body['startDateTime'] = gmdate('Y-m-d\\TH:i:s\\Z', $startTimestamp);
            $effectiveEnd = $endTimestamp ?: ($startTimestamp + 7200);
            $body['endDateTime'] = gmdate('Y-m-d\\TH:i:s\\Z', $effectiveEnd);
        } else if (!empty($endTimestamp)) {
            $body['endDateTime'] = gmdate('Y-m-d\\TH:i:s\\Z', $endTimestamp);
        }

        if ($recordAutomatically) {
            $body['recordAutomatically'] = true;
        }

        if (!empty($coOrganizerIds)) {
            $body['participants']['coOrganizers'] = array_map(static fn(string $id) => [
                'identity' => [
                    'user' => ['id' => $id],
                ],
            ], $coOrganizerIds);
        }

        if (!empty($roomEmail)) {
            $body['participants']['attendees'][] = [
                'upn' => $roomEmail,
                'role' => 'attendee',
            ];
        }

        $response = $this->legacyRequest('POST', "users/$userId/onlineMeetings", $body, 'beta');
        $data = $this->graphResponseToArray($response);
        return $this->arrayToOnlineMeeting($data);
    }

    public function getAttendanceRecords(string $organizerId, string $meetingId): array
    {
        $path = "users/$organizerId/onlineMeetings/$meetingId/attendanceReports?\$expand=attendanceRecords";
        $response = $this->legacyRequest('GET', $path, [], 'beta');
        $data = $this->graphResponseToArray($response);

        $records = [];
        foreach (($data['value'] ?? []) as $report) {
            if (!empty($report['attendanceRecords']) && is_array($report['attendanceRecords'])) {
                $records = array_merge($records, $report['attendanceRecords']);
            }
        }

        return $records;
    }

    public function getOnlineMeetingByVideoTeleconferenceId(string $meetingId): mixed
    {
        $query = http_build_query(['$filter' => "VideoTeleconferenceId eq '$meetingId'"]);
        $response = $this->legacyRequest('GET', "communications/onlineMeetings?$query", [], 'v1.0');
        $data = $this->graphResponseToArray($response);
        $first = $data['value'][0] ?? null;

        if (empty($first) || !is_array($first)) {
            return null;
        }

        return $this->arrayToOnlineMeeting($first);
    }

    public function deleteOnlineMeetingById(string $meetingId): mixed
    {
        return $this->legacyRequest('DELETE', "communications/onlineMeetings/$meetingId", [], 'v1.0');
    }

    public function getTeamsRooms(): array
    {
        $rooms = [];
        $query = http_build_query([
            '$select' => 'displayName,emailAddress',
            '$top' => 999,
        ]);

        try {
            $response = $this->legacyRequest('GET', "places/microsoft.graph.room?$query", [], 'v1.0');
            $data = $this->graphResponseToArray($response);
        } catch (\Throwable $e) {
            $response = $this->legacyRequest('GET', "places/microsoft.graph.room?$query", [], 'beta');
            $data = $this->graphResponseToArray($response);
        }

        foreach (($data['value'] ?? []) as $room) {
            $email = $room['emailAddress'] ?? null;
            $name = $room['displayName'] ?? $email;
            if (!empty($email)) {
                $rooms[(string)$email] = (string)$name;
            }
        }

        return $rooms;
    }

    private function normalizeGraphPath(string $url): string
    {
        if ($url === '') {
            return '/';
        }
        if (str_starts_with($url, 'http://') || str_starts_with($url, 'https://')) {
            return $url;
        }
        return str_starts_with($url, '/') ? $url : '/' . $url;
    }

    private function graphResponseToArray(mixed $response): array
    {
        if (is_array($response)) {
            return $response;
        }

        if (is_object($response) && method_exists($response, 'getBody')) {
            $body = $response->getBody();
            if (is_array($body)) {
                return $body;
            }
            if (is_object($body)) {
                return json_decode(json_encode($body), true) ?? [];
            }
        }

        return json_decode(json_encode($response), true) ?? [];
    }

    private function arrayToOnlineMeeting(array $data): OnlineMeeting
    {
        $meeting = new OnlineMeeting();

        if (isset($data['id']) && method_exists($meeting, 'setId')) {
            $meeting->setId((string)$data['id']);
        }
        if (isset($data['joinWebUrl']) && method_exists($meeting, 'setJoinWebUrl')) {
            $meeting->setJoinWebUrl((string)$data['joinWebUrl']);
        }
        if (isset($data['subject']) && method_exists($meeting, 'setSubject')) {
            $meeting->setSubject((string)$data['subject']);
        }
        if (isset($data['startDateTime']) && method_exists($meeting, 'setStartDateTime')) {
            $meeting->setStartDateTime((string)$data['startDateTime']);
        }
        if (isset($data['endDateTime']) && method_exists($meeting, 'setEndDateTime')) {
            $meeting->setEndDateTime((string)$data['endDateTime']);
        }

        $meeting->setAdditionalData($data);
        return $meeting;
    }

    private function getLegacyAccessToken(): string
    {
        if (!empty($this->legacyAccessToken) && $this->legacyAccessTokenExpireAt > (time() + 30)) {
            return $this->legacyAccessToken;
        }

        $config = [
            'verify' => true,
            'timeout' => 30,
        ];

        if (!empty($this->proxyUrl)) {
            $config['proxy'] = [
                'http' => $this->proxyUrl,
                'https' => $this->proxyUrl,
            ];
        }

        $provider = new GenericProvider([
            'clientId' => $this->clientId,
            'clientSecret' => $this->clientSecret,
            'redirectUri' => '',
            'urlAuthorize' => "https://login.microsoftonline.com/{$this->tenantId}/oauth2/v2.0/authorize",
            'urlAccessToken' => "https://login.microsoftonline.com/{$this->tenantId}/oauth2/v2.0/token",
            'urlResourceOwnerDetails' => 'https://graph.microsoft.com/oidc/userinfo',
            'scopes' => 'https://graph.microsoft.com/.default',
            'verify' => true,
            'timeout' => 30,
        ], [
            'httpClient' => new GuzzleClient($config),
        ]);

        $token = $provider->getAccessToken('client_credentials', [
            'scope' => 'https://graph.microsoft.com/.default',
        ]);

        $this->legacyAccessToken = $token->getToken();
        $this->legacyAccessTokenExpireAt = (int)($token->getExpires() ?? (time() + 3000));
        return $this->legacyAccessToken;
    }
}
