<?php

namespace BCedric\UCAOffice365\Service;

use DateTime;
use Exception;
use Microsoft\Graph\GraphServiceClient;
use Microsoft\Kiota\Abstractions\ApiException;
use Microsoft\Kiota\Authentication\Oauth\ClientCredentialContext;
use Microsoft\Graph\Generated\Models\AssignedLicense;
use Microsoft\Graph\Generated\Models\Entity;
use Microsoft\Graph\Generated\Models\Group;
use Microsoft\Graph\Generated\Models\Identity;
use Microsoft\Graph\Generated\Models\IdentitySet;
use Microsoft\Graph\Generated\Models\LicenseDetails;
use Microsoft\Graph\Generated\Models\MeetingParticipantInfo;
use Microsoft\Graph\Generated\Models\MeetingParticipants;
use Microsoft\Graph\Generated\Models\OnlineMeeting;
use Microsoft\Graph\Generated\Models\User;
use Microsoft\Graph\Generated\Models\UserSettings;
use Symfony\Component\DependencyInjection\Attribute\Autowire;
use Symfony\Contracts\HttpClient\HttpClientInterface;
use Microsoft\Graph\Generated\Drives\Item\Items\Item\Children\ChildrenRequestBuilderGetRequestConfiguration;
use Microsoft\Graph\Generated\Drives\Item\Items\Item\Children\ChildrenRequestBuilderGetQueryParameters;

class GraphAPI
{
    private ?GraphServiceClient $graphServiceClient = null;
    public $userUrlPrefix = "https://graph.microsoft.com/v1.0/users/";
    private $httpClient;

    public function __construct(
        #[Autowire(env: 'GRAPH_TENANT')] private readonly string $tenantId,
        #[Autowire(env: 'GRAPH_CLIENT')] private readonly string $clientId,
        #[Autowire(env: 'GRAPH_CLIENT_SECRET')] private readonly string $clientSecret,
        private readonly HttpClientInterface $client
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
            $this->graphServiceClient = new GraphServiceClient($tokenRequestContext);
        }

        return $this->graphServiceClient;
    }

    /**
     * @deprecated
     */
    public function getGraphApi()
    {
        return $this->getGraphServiceClient();
    }

    public function getUser($email) : User
    {
        try {
            $graphServiceClient = $this->getGraphServiceClient();
            try {
                $user = $graphServiceClient->users()->byUserId($email)->get()->wait();
                return $user;
            } catch (ApiException $ex) {
               throw $ex;
            }
        } catch (ApiException $exception) {
            throw $exception;
        }
    }


    public function getSharepointTeam(string $userId) : \Microsoft\Graph\Generated\Models\DirectoryObjectCollectionResponse
    {
        try {
            $graphServiceClient = $this->getGraphServiceClient();

            $requestConfiguration = new \Microsoft\Graph\Generated\Users\Item\OwnedObjects\OwnedObjectsRequestBuilderGetRequestConfiguration();
            $requestConfiguration->queryParameters = new \Microsoft\Graph\Generated\Users\Item\OwnedObjects\OwnedObjectsRequestBuilderGetQueryParameters();
            $requestConfiguration->queryParameters->select = ['id', 'displayName'];

            $ownedObjects = $graphServiceClient->users()->byUserId($userId)->ownedObjects()->get($requestConfiguration)->wait();
            return $ownedObjects;
        } catch (ApiException $exception) {
            throw $exception;
        }
    }

    public function getArrayOfTeams($sharePointList) : array
    {
        $teamsList = [];
        if(!empty($sharePointList)) {
            foreach ($sharePointList->getValue() as $sharepoint) {
                $teamsList[] = array('displayName' => $sharepoint->getDisplayName(), 'id' => $sharepoint->getId());
            }
        }
        return $teamsList;
    }

    public function getPersonnalDriveId($userId)
    {
        $client = $this->getGraphServiceClient();
        try {
            $drive = $client->users()->byUserId($userId)->drive()->get()->wait();
            $driveId = $drive->getId();
            return $driveId;
        } catch (\Throwable $e) {
            return new Exception($e);
        }
    }

    public function getDriveId($groupId)
    {
        $client = $this->getGraphServiceClient();
        try {
            $drive = $client->groups()->byGroupId($groupId)->drive()->get()->wait();
            $driveId = $drive->getId();
        } catch (\Throwable $e) {
            return new Exception($e);
        }

        return $driveId;
    }


    public function getSharepointDriveVideos(string $driveId): array| Exception
    {
        $client = $this->getGraphServiceClient();
        try {
            return $this->browseDrive($client, $driveId);
        } catch (\Throwable $e) {
            return new Exception($e);
        }

    }

    private function browseDrive(
        \Microsoft\Graph\GraphServiceClient $client,
        string $driveId,
    ): array {
        $files = [];
        $items = $client
            ->drives()
            ->byDriveId($driveId)
            ->items()
            ->byDriveItemId('root')
            ->searchWithQ('.mp4')
            ->get()
            ->wait();

        foreach ($items->getValue() as $item) {
            $files[] = array('shareId' => $driveId, 'mediaName' => $item->getName(), 'mediaDate' => $item->getCreatedDateTime(), 'mediaId' => $item->getId(), 'mediaURL' => $item->getWebUrl());
        }

        return $files;
    }

    public function getMediaContentFromDrive(string $driveId, string $mediaId)
    {
        $client = $this->getGraphServiceClient();
        try {
            $driveItem = $client->drives()->byDriveId($driveId)->items()->byDriveItemId($mediaId)->get()->wait();
            // Download Url not the WebUrl
            $additionalData = $driveItem->getAdditionalData();
            if (isset($additionalData['@microsoft.graph.downloadUrl'])) {
                return $additionalData['@microsoft.graph.downloadUrl'];
            }

        } catch (\Throwable $e) {
            return new Exception($e);
        }
    }

    /** DEPRECATED **/

    public function getSharepointOwner($groupId)
    {
        $url = "/groups/" . $groupId . "/owners";
        try {
            $graph = $this->getGraphApi();
            return $graph->createRequest("GET", $url)->execute();
        } catch (Exception $exception) {
            throw $exception;
        }
    }

    public function getUserGroups($email)
    {
        $url = "/users/$email/memberOf";
        try {
            $graph = $this->getGraphApi();
            return $graph->createRequest("GET", $url)->execute();
        } catch (Exception $exception) {
            throw $exception;
        }
    }

    /**
     * @param string $email
     * @return string
     */
    public function getUserId(string $email)
    {
        global $CFG;
        $queryParams = array(
            '$filter' => "userPrincipalName eq '$email' or mail eq '$email'",
        );
        //        $url = '/users/'.$email;
        $url = '/users?' . http_build_query($queryParams);

        try {
            $graph = $this->getGraphApi();
            $user = $graph->createRequest("GET", $url)
                ->setReturnType(User::class)

                ->execute();

            if (count($user) > 1) {
                throw new Exception("The email is not unique in Azure AD.");
            }

            //            return ($user) ? $user->getId() : null;
            return ($user[0]) ? $user[0]->getId() : null;
        } catch (Exception $exception) {
            // error_log("[" . date("Y-m-d H:i:s") . "] " . $email . " - ERROR: " . $exception->getResponse()->getBody()->getContents() . "\n", 3, $CFG->dataroot . '/clfd/ucateams.log');
            throw $exception;
        }
    }

    /**
     * @param string $groupId
     * @param array $memberSettings
     * @return mixed
     * @throws Exception
     */
    public function createTeam(string $groupId, array $memberSettings = [])
    {
        $graph = $this->getGraphApi();

        $parameters = ["memberSettings" => $memberSettings];
        if (empty($parameters["memberSettings"])) {
            $parameters["memberSettings"] = [
                "allowCreateUpdateChannels" => false,
                "allowDeleteChannels" => false,
                "allowAddRemoveApps" => false,
                "allowCreateUpdateRemoveTabs" => false,
                "allowCreateUpdateRemoveConnectors" => false,
            ];
        }

        try {
            return $graph->createRequest("PUT", "/groups/$groupId/team")
                ->setReturnType(Group::class)
                ->attachBody(json_encode($parameters))
                ->execute();
        } catch (Exception $exception) {
            throw $exception;
        }
    }

    /**
     * @param string $name
     * @param string $description
     * @param array|string $ownersId
     * @return string
     * @throws Exception
     */
    public function createGroup(string $name, string $description, $ownersId)
    {
        $users = [];
        foreach ((array) $ownersId as $ownerId) {
            $users[] = $this->userUrlPrefix . $ownerId;
        }

        $group = new Group();
        $group->setDisplayName($name);
        $group->setMailNickname(preg_replace("/[^A-Za-z0-9]/", '', $name) . uniqid());
        $group->setDescription($description);
        $group->setVisibility("Private");
        $group->setGroupTypes(["Unified"]);
        $group->setMailEnabled(true);
        $group->setSecurityEnabled(false);
        $group->setOwners($users);
        $group->setMembers($users);

        $data = $group->jsonSerialize();

        $data["owners@odata.bind"] = $data["owners"];
        $data["members@odata.bind"] = $data["members"];
        $data["resourceBehaviorOptions"] = ["WelcomeEmailDisabled"];

        unset($data["owners"]);
        unset($data["members"]);

        $graph = $this->getGraphApi();

        try {
            $response = $graph->createRequest("POST", "/groups")
                ->attachBody($data)
                ->setReturnType(Group::class)
                ->execute();
            return $response->getId();
        } catch (Exception $exception) {
            throw $exception;
        }
    }

    /**
     * @param $groupId
     * @return mixed
     * @throws \Microsoft\Graph\Exception\GraphException
     */
    public function readTeam($groupId)
    {
        $url = "/groups/$groupId/team";
        $graph = $this->getGraphApi();
        return $graph->createRequest("GET", $url)
            ->setReturnType(Group::class)
            ->execute();
    }

    /**
     * @param string $groupId
     * @param array $memberSettings
     * @return mixed
     * @throws Exception
     */
    public function copyTeam(string $team, string $name, string $description, $ownersId)
    {
        $graph = $this->getGraphApi();
        $users = [];
        foreach ((array) $ownersId as $ownerId) {
            $users[] = $this->userUrlPrefix . $ownerId;
        }

        $group = new Group();
        $group->setDisplayName($name);
        $group->setMailNickname(preg_replace("/[^A-Za-z0-9]/", '', $name) . uniqid());
        $group->setDescription($description);
        $group->setVisibility("Private");
        //        $group->setGroupTypes(["Unified"]);
        //        $group->setMailEnabled(true);
        //        $group->setSecurityEnabled(false);
        $group->setOwners($users);
        //        $group->setMembers($users);

        $data = $group->jsonSerialize();
        $data["partsToClone"] = "apps,tabs,settings,channels";
        $data["resourceBehaviorOptions"] = ["WelcomeEmailDisabled"];
        //        $data["owners"]=$users;
        //        $data["owners@odata.bind"]=$users;
        //        $data["owners@odata.bind"]=$data["owners"];

        unset($data["owners"]);
        //        unset($data["members"]);

        try {
            $response = $graph->createRequest("POST", "/teams/$team/clone")
                //                ->setReturnType(\Microsoft\Graph\Http\GraphResponse::class)
                ->attachBody($data)
                ->execute();

            if ($response) {
                $location = $response->getHeaders()["Location"][0];
                $team_id = explode("'", explode("/", $location)[1])[1];
                $team = $this->readTeam($team_id);

                return $team;
            }
        } catch (Exception $exception) {
            throw $exception;
        }
    }

    /**
     * @param array|string $ownersId
     * @return string
     * @throws Exception
     */
    public function updateGroupOwners($group, $ownersId)
    {
        global $USER;
        $current = $this->getUserId($USER->email);
        $graph = $this->getGraphApi();

        try {
            $response = $this->addOwner($current, $group, false);
            if ($response && $response->getStatus() == 204) {
                //ok
                $this->deleteModelOwners($group, false, $current);
            }
        } catch (Exception $exception) {
            throw $exception;
        }
    }

    /**
     * @param $ownerid
     * @param $groupid
     * @param bool $retry
     * @return mixed
     */
    public function addOwner($ownerid, $groupid, $retry = false)
    {
        $graph = $this->getGraphApi();
        $data = json_encode(["@odata.id" => $this->userUrlPrefix . $ownerid]);

        try {
            return $graph->createRequest("POST", "/groups/" . $groupid . "/owners/\$ref")
                ->attachBody($data)
                ->execute();
        } catch (Exception $e) {
            if (!$retry && $e->getCode() == 404) {
                //Si on est tombé sur une 404, on retente une fois pour pas que l'action soit trop longue
                sleep(10);
                return $this->addOwner($ownerid, $groupid, true);
            }

            return null;
        }
    }

    /**
     * @param $groupid
     * @param bool $retry
     * @param null $current
     * @throws Exception
     */
    public function deleteModelOwners($groupid, $retry = false, $current = null)
    {
        global $USER;
        $current = ($current) ? $current : $this->getUserId($USER->email);
        $graph = $this->getGraphApi();

        try {
            $response = $graph->createRequest("GET", "/groups/" . get_config('ucateams', 'team_model') . "/owners")
                ->execute();
            $model_owners = [];
            foreach ($response->getBody()["value"] as $owner) {
                $model_owners[] = $owner['id'];
            }

            foreach ($model_owners as $owner) {
                if ($owner != $current) {
                    $graph->createRequest("DELETE", "/groups/" . $groupid . "/owners/" . $owner . "/\$ref")
                        ->execute();
                }
            }

            return true;
        } catch (Exception $e) {
            return false;
        }
    }

    /**
     * @param string $userId
     * @param string $subject
     * @param DateTime $startDateTime
     * @param DateTime $endDateTime
     * @return string
     * @throws \Microsoft\Graph\Exception\GraphException
     */
    public function createBroadcastEvent(string $userId, string $subject, DateTime $startDateTime, DateTime $endDateTime)
    {
        $onlineMeeting = new OnlineMeeting();
        $onlineMeeting->setStartDateTime($startDateTime->format("Y-m-d\TH:i:s\Z"));
        $onlineMeeting->setEndDateTime($endDateTime->format("Y-m-d\TH:i:s\Z"));
        /*
        $onlineMeeting->setStartDateTime($startDate->getTimestamp()+11644473600);
        $onlineMeeting->setEndDateTime($endDate->getTimestamp()+11644473600);
         */
        $onlineMeeting->setSubject($subject);
        $user = new Identity();
        $user->setId($userId);
        $identity = new IdentitySet();
        $identity->setUser($user->jsonSerialize());
        $participant = new MeetingParticipantInfo();
        $participant->setIdentity($identity->jsonSerialize());
        $participants = new MeetingParticipants();
        $participants->setOrganizer($participant->jsonSerialize());
        $onlineMeeting->setParticipants($participants->jsonSerialize());
        $data = $onlineMeeting->jsonSerialize();
        $lobbyBypassSettings = new Entity(["scope" => "everyone", "isDialInBypassEnabled" => true]);
        $data["lobbyBypassSettings"] = $lobbyBypassSettings->jsonSerialize();
        $data["autoAdmittedUsers"] = "everyone";
        $data["allowedPresenters"] = "roleIsPresenter";
        $graph = $this->getGraphApi("beta");
        try {
            $response = $graph->createRequest("POST", "/communications/onlineMeetings")
                ->attachBody($data)
                ->setReturnType(OnlineMeeting::class)
                ->execute();
            return $response;
        } catch (Exception $exception) {
            throw $exception;
        }
    }

    public function getBroadcastEvent(string $meetingId)
    {
        $queryParams = array(
            '$filter' => "VideoTeleconferenceId eq '$meetingId'",
        );
        $url = '/communications/onlineMeetings/?' . http_build_query($queryParams);
        $graph = $this->getGraphApi();
        return $graph->createRequest("GET", $url)
            ->setReturnType(OnlineMeeting::class)
            ->execute();
    }
}
