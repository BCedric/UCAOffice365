<?php

namespace BCedric\UCAOffice365\Service;

use Exception;
use Microsoft\Graph\Generated\Communications\OnlineMeetings\OnlineMeetingsRequestBuilderGetRequestConfiguration;
use Microsoft\Graph\Generated\Models\ReferenceCreate;
use Microsoft\Graph\Generated\Models\OnlineMeeting;
use Microsoft\Graph\Generated\Places\GraphRoom\GraphRoomRequestBuilderGetRequestConfiguration;
use Microsoft\Graph\GraphServiceClient;
use Symfony\Component\DependencyInjection\Attribute\Autowire;

class GraphAPITeams
{
    public function __construct(
        private readonly GraphAPI $graphAPI,
        #[Autowire(env: 'GRAPH_TENANT')] private readonly string $tenantId,
        #[Autowire(env: 'GRAPH_CLIENT')] private readonly string $clientId,
        #[Autowire(env: 'GRAPH_CLIENT_SECRET')] private readonly string $clientSecret,
        #[Autowire(env: 'PROXY_URL')] private readonly ?string $proxyUrl = null,
    ) {
    }

    public function getGraphServiceClient(): GraphServiceClient
    {
        return $this->graphAPI->getGraphServiceClient();
    }

    public function legacyRequest(
        string $method,
        string $url,
        array $body = [],
        string $version = 'v1.0',
        ?string $returntype = null
    ): mixed {
        $data = $this->graphAPI->rawJsonRequest($method, $url, !empty($body) ? $body : null, $version);
        if ($returntype === OnlineMeeting::class && is_array($data)) {
            return $this->arrayToOnlineMeeting($data);
        }
        return $data;
    }

    public function addGroupMemberByUserId(string $groupId, string $userId): mixed
    {
        $client = $this->getGraphServiceClient();
        $reference = new ReferenceCreate();
        $reference->setOdataId($this->graphAPI->userUrlPrefix . $userId);

        try {
            return $client
                ->groups()
                ->byGroupId($groupId)
                ->members()
                ->ref()
                ->post($reference)
                ->wait();
        } catch (\Throwable $e) {
            throw new Exception($e->getMessage(), (int)$e->getCode(), $e);
        }
    }

    public function getUserId(string $email): ?string
    {
        return $this->graphAPI->getUserId($email);
    }

    public function createTeam(string $groupId, array $memberSettings = []): mixed
    {
        return $this->graphAPI->createTeam($groupId, $memberSettings);
    }

    public function createGroup(string $name, string $description, array|string $ownersId): ?string
    {
        return $this->graphAPI->createGroup($name, $description, $ownersId);
    }

    public function readTeam(string $groupId): mixed
    {
        return $this->graphAPI->readTeam($groupId);
    }

    public function copyTeam(string $team, string $name, string $description, array|string $ownersId): mixed
    {
        return $this->graphAPI->copyTeam($team, $name, $description, $ownersId);
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
            $path = "/groups/$groupId/members?\$select=id&\$top=999";

            while (!empty($path)) {
                $data = $this->graphAPI->rawJsonRequest('GET', $path);

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
        $client = $this->getGraphServiceClient();
        $meeting = new OnlineMeeting();
        $meeting->setSubject($subject);
        $additionalData = [];

        if (!empty($startTimestamp)) {
            $meeting->setStartDateTime(gmdate('Y-m-d\\TH:i:s\\Z', $startTimestamp));
            $effectiveEnd = $endTimestamp ?: ($startTimestamp + 7200);
            $meeting->setEndDateTime(gmdate('Y-m-d\\TH:i:s\\Z', $effectiveEnd));
        } else if (!empty($endTimestamp)) {
            $meeting->setEndDateTime(gmdate('Y-m-d\\TH:i:s\\Z', $endTimestamp));
        }

        if ($recordAutomatically) {
            $additionalData['recordAutomatically'] = true;
        }

        if (!empty($coOrganizerIds)) {
            $additionalData['participants']['coOrganizers'] = array_map(static fn(string $id) => [
                'identity' => [
                    'user' => ['id' => $id],
                ],
            ], $coOrganizerIds);
        }

        if (!empty($roomEmail)) {
            $additionalData['participants']['attendees'][] = [
                'upn' => $roomEmail,
                'role' => 'attendee',
            ];
        }

        if (!empty($additionalData)) {
            $meeting->setAdditionalData($additionalData);
        }

        return $client
            ->users()
            ->byUserId($userId)
            ->onlineMeetings()
            ->post($meeting)
            ->wait();
    }

    public function getAttendanceRecords(string $organizerId, string $meetingId): array
    {
        $client = $this->getGraphServiceClient();

        $records = [];
        try {
            $reports = $client
                ->users()
                ->byUserId($organizerId)
                ->onlineMeetings()
                ->byOnlineMeetingId($meetingId)
                ->attendanceReports()
                ->get()
                ->wait();

            if (!empty($reports) && method_exists($reports, 'getValue')) {
                foreach (($reports->getValue() ?? []) as $report) {
                    $attendanceRecords = method_exists($report, 'getAttendanceRecords') ? $report->getAttendanceRecords() : [];
                    if (!empty($attendanceRecords)) {
                        $records = array_merge($records, json_decode(json_encode($attendanceRecords), true) ?? []);
                    }
                }
            }

            return $records;
        } catch (\Throwable $e) {
            $path = "users/$organizerId/onlineMeetings/$meetingId/attendanceReports?\$expand=attendanceRecords";
            $response = $this->legacyRequest('GET', $path, [], 'v1.0');
            $data = $this->graphResponseToArray($response);

            foreach (($data['value'] ?? []) as $report) {
                if (!empty($report['attendanceRecords']) && is_array($report['attendanceRecords'])) {
                    $records = array_merge($records, $report['attendanceRecords']);
                }
            }

            return $records;
        }
    }

    public function getOnlineMeetingByVideoTeleconferenceId(string $meetingId): mixed
    {
        $client = $this->getGraphServiceClient();
        $requestConfiguration = new OnlineMeetingsRequestBuilderGetRequestConfiguration();
        $requestConfiguration->queryParameters = OnlineMeetingsRequestBuilderGetRequestConfiguration::createQueryParameters(
            filter: "VideoTeleconferenceId eq '$meetingId'"
        );

        $meetings = $client
            ->communications()
            ->onlineMeetings()
            ->get($requestConfiguration)
            ->wait();

        if (empty($meetings) || !method_exists($meetings, 'getValue')) {
            return null;
        }

        $first = ($meetings->getValue() ?? [])[0] ?? null;
        return $first ?: null;
    }

    public function deleteOnlineMeetingById(string $meetingId): mixed
    {
        return $this->getGraphServiceClient()
            ->communications()
            ->onlineMeetings()
            ->byOnlineMeetingId($meetingId)
            ->delete()
            ->wait();
    }

    public function getTeamsRooms(): array
    {
        $rooms = [];
        $requestConfiguration = new GraphRoomRequestBuilderGetRequestConfiguration();
        $requestConfiguration->queryParameters = GraphRoomRequestBuilderGetRequestConfiguration::createQueryParameters(
            select: ['displayName', 'emailAddress'],
            top: 999
        );

        $response = $this->getGraphServiceClient()
            ->places()
            ->graphRoom()
            ->get($requestConfiguration)
            ->wait();

        foreach (($response?->getValue() ?? []) as $room) {
            $email = method_exists($room, 'getEmailAddress') ? $room->getEmailAddress() : null;
            $name = method_exists($room, 'getDisplayName') ? $room->getDisplayName() : $email;
            if (!empty($email)) {
                $rooms[(string)$email] = (string)$name;
            }
        }

        return $rooms;
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
}
