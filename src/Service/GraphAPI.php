<?php

namespace BCedric\UCAOffice365\Service;

use DateTime;
use Exception;
use Microsoft\Graph\Graph;
use Microsoft\Graph\Model\AssignedLicense;
use Microsoft\Graph\Model\Entity;
use Microsoft\Graph\Model\Group;
use Microsoft\Graph\Model\Identity;
use Microsoft\Graph\Model\IdentitySet;
use Microsoft\Graph\Model\LicenseDetails;
use Microsoft\Graph\Model\MeetingParticipantInfo;
use Microsoft\Graph\Model\MeetingParticipants;
use Microsoft\Graph\Model\OnlineMeeting;
use Microsoft\Graph\Model\User;
use Microsoft\Graph\Model\UserSettings;
use Symfony\Component\DependencyInjection\Attribute\Autowire;
use Symfony\Contracts\HttpClient\HttpClientInterface;

class GraphAPI
{
    private $tenantId;
    private $clientId;
    private $clientSecret;
    private $token;
    public $userUrlPrefix = "https://graph.microsoft.com/v1.0/users/";
    private $httpClient;

    public function __construct(
        #[Autowire(env: 'GRAPH_TENANT')] string $tenantId,
        #[Autowire(env: 'GRAPH_CLIENT')] string $clientId,
        #[Autowire(env: 'GRAPH_CLIENT_SECRET')] string $clientSecret,
        private readonly HttpClientInterface $client
    ) {
        $this->tenantId = $tenantId;
        $this->clientId = $clientId;
        $this->clientSecret = $clientSecret;
    }

    public function getToken()
    {
        if (is_null($this->token)) {
            $this->token = $this->generateToken();
        } elseif ($this->token->expires_on <= time()) {
            $this->token = $this->generateToken();
        }

        return $this->token;
    }

    /**
     * @return mixed
     * @throws Exception
     */
    public function generateToken()
    {
        $url = 'https://login.microsoftonline.com/' . $this->tenantId . '/oauth2/token?api-version=1.0';

        try {
            $options = [
                'body' => [
                    'client_id' => $this->clientId,
                    'client_secret' => $this->clientSecret,
                    'resource' => 'https://graph.microsoft.com/',
                    'grant_type' => 'client_credentials',
                ],
            ];

            if (isset($_ENV['PROXY_URL'])) {
                $options['proxy'] = $_ENV['PROXY_URL'];
            }

            $response = $this->httpClient->request('POST', $url, $options);
            $token = json_decode($response->getContent());
            return $token;
        } catch (Exception $exception) {
            throw $exception;
        }
    }

    public function getGraphApi($version = "v1.0")
    {
        $graph = new Graph();
        if (isset($_ENV['PROXY_URL'])) {
            $graph->setProxyPort($_ENV['PROXY_URL']);
        }

        $graph->setApiVersion($version);
        $graph->setAccessToken($this->getToken()->access_token);
        return $graph;
    }

    public function getUser($email)
    {
        $queryParams = array(
            '$filter' => "userPrincipalName eq '$email' or mail eq '$email'",
        );
        $url = '/users?' . http_build_query($queryParams);
        try {
            $graph = $this->getGraphApi();
            return $graph->createRequest("GET", $url)->setReturnType(User::class)->execute();
        } catch (Exception $exception) {
            throw $exception;
        }
    }

    public function getUserSettings($email)
    {
        $queryParams = array(
            '$filter' => "userPrincipalName eq '$email' or mail eq '$email'",
        );
        $url = '/users/' . $this->getUserId($email) . '/settings';
        try {
            $graph = $this->getGraphApi();
            return $graph->createRequest("GET", $url)->setReturnType(UserSettings::class)->execute();
        } catch (Exception $exception) {
            throw $exception;
        }
    }

    public function getUserMailboxSettings($email)
    {
        $url = '/users/' . $email . '/mailboxSettings';
        try {
            $graph = $this->getGraphApi();
            return $graph->createRequest("GET", $url)->setReturnType(UserSettings::class)->execute();
        } catch (Exception $exception) {
            throw $exception;
        }
    }

    public function getUserLicenseDetails($userId)
    {
        $url = '/users/' . $userId . '/licenseDetails';
        try {
            $graph = $this->getGraphApi();
            return $graph->createRequest("GET", $url)->setReturnType(LicenseDetails::class)->execute();
        } catch (Exception $exception) {
            throw $exception;
        }
    }

    public function getUserCalendars($email, $limitDate)
    {

        $url = '/users/' . $email . "/events?" . '$filter' . "=start/dateTime ge '$limitDate'";

        try {
            $graph = $this->getGraphApi();
            $response = $graph->createRequest("GET", $url . '&$count=true')->execute();
            $resBody = $response->getBody();
            $count = $resBody['@odata.count'];
            return $graph->createRequest("GET", $url . '&$top=' . $count)->execute();
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

    public function getUserAssignedLicenses($email)
    {
        $queryParams = array(
            '$filter' => "userPrincipalName eq '$email' or mail eq '$email'",
        );
        $url = '/users/' . $this->getUserId($email) . '/assignedLicenses';
        try {
            $graph = $this->getGraphApi();
            return $graph->createRequest("GET", $url)->setReturnType(AssignedLicense::class)->execute();
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

    public function getSharepointTeam(string $userId)
    {
        $select = '$select';
        $url = "/users/" . $userId . "/ownedObjects?$select=id,displayName";
        try {
            $graph = $this->getGraphApi();
            return $graph->createRequest("GET", $url)->execute();
        } catch (Exception $exception) {
            throw $exception;
        }
    }

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

    public function getSharepointDrive(string $sharepointId)
    {
        $select = '$select';
        $url = "/groups/" . $sharepointId . "/drive/root/search(q='.mp4')?$select=id,name,createdDateTime,webUrl";
        try {
            $graph = $this->getGraphApi();
            return $graph->createRequest("GET", $url)->execute();
        } catch (Exception $exception) {
            throw $exception;
        }
    }

    public function getPersonnalSharepointDrive(string $user)
    {
        $select = '$select';
        $url = "/users/" . $user . "/drive/root/search(q='.mp4')?$select=id,name,createdDateTime,webUrl";
        try {
            $graph = $this->getGraphApi();
            return $graph->createRequest("GET", $url)->execute();
        } catch (Exception $exception) {
            throw $exception;
        }
    }

    public function getSharepointDriveId(string $sharepointId)
    {
        $url = "/groups/" . $sharepointId . "/drive/root/";
        try {
            $graph = $this->getGraphApi();
            $driveInfo = $graph->createRequest("GET", $url)->execute();
            return $driveInfo->getBody()['parentReference']['driveId'];
        } catch (Exception $exception) {
            throw $exception;
        }
    }

    public function getPersonnalSharepointDriveId(string $userId)
    {
        $url = "/users/" . $userId . "/drive/root/";
        try {
            $graph = $this->getGraphApi();
            $driveInfo = $graph->createRequest("GET", $url)->execute();
            return $driveInfo->getBody()['parentReference']['driveId'];
        } catch (Exception $exception) {
            throw $exception;
        }
    }

    public function getMediaContentFromDrive(string $driveId, string $mediaId)
    {
        $url = "/drives/" . $driveId . "/items/" . $mediaId;
        try {
            $graph = $this->getGraphApi();
            $driveInfo = $graph->createRequest("GET", $url)->execute();
            return $driveInfo->getBody()['@microsoft.graph.downloadUrl'];
        } catch (Exception $exception) {
            throw $exception;
        }
    }

    public function getMediaUrlFromDrive(string $driveId, string $mediaId)
    {
        $url = "/drives/" . $driveId . "/items/" . $mediaId;
        try {
            $graph = $this->getGraphApi();
            $driveInfo = $graph->createRequest("GET", $url)->execute();
            return $driveInfo->getBody()['webUrl'];
        } catch (Exception $exception) {
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
