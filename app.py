"""App to sort emails by the user's priorities and deadlines"""
import concurrent.futures
import datetime
import json
import os
import signal
import sys
import urllib.parse
from typing import Any, NoReturn

import requests
from dotenv import load_dotenv
from flask import Flask, render_template, request, redirect
from google import genai

from util import diskcache, fetch, writefile, html_to_pdf

load_dotenv()
app = Flask(__name__)
genai_client = genai.Client()  # TODO: if lite no enuf 4 dis, change here vvv
chat = genai_client.chats.create(model='gemini-2.5-flash')  # meow!


class ApiError(Exception):
    def __init__(self, o: object):
        super().__init__(json.dumps(o, indent=2))


def init_office():
    fetch("https://outlook.office.com/owa/startupdata.ashx?app=Mail&n=0", {
        "credentials": "include",
        "headers": {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:143.0) Gecko/20100101 Firefox/143.0",
            "Accept": "*/*",
            "Accept-Language": "en-GB,en;q=0.5",
            "action": "StartupData",
            "calendarviewparams": "{\"TimeZoneStr\":\"GMT Standard Time\",\"RangeStart\":\"2025-09-22T00:00:00+01:00\",\"RangeEnd\":\"2025-09-29T00:00:00+01:00\",\"FolderId\":{\"__type\":\"FolderId:#Exchange\",\"Id\":\"AQMkADAzODA5NzA3LWI3NTEtNDAxMC04MmMwLWNjADU4MDQ4MjYyODIALgAAA/uV1hy/6XhAtrtIgOrOGv4BAI2thI0NWwZPqja/whIfpMwAAAIBDQAAAA==\",\"ChangeKey\":\"AgAAAA==\"}}",
            "folderparams": "{\"TimeZoneStr\":\"GMT Standard Time\",\"FolderPaneBitFlags\":2}",
            "messageparams": "{\"TimeZoneStr\":\"GMT Standard Time\",\"InboxReadingPanePosition\":1,\"IsFocusedInboxOn\":false,\"BootWithConversationView\":true,\"SortResults\":[{\"Path\":{\"__type\":\"PropertyUri:#Exchange\",\"FieldURI\":\"conversation:LastDeliveryOrRenewTime\"},\"Order\":\"Descending\"},{\"Path\":{\"__type\":\"PropertyUri:#Exchange\",\"FieldURI\":\"conversation:LastDeliveryTime\"},\"Order\":\"Descending\"}],\"IsSenderScreeningSettingEnabled\":false}",
            "ms-cv": os.getenv('OWA_INIT_MS_CV'),
            "prefer": "exchange.behavior=\"IncludeThirdPartyOnlineMeetingProviders\"",
            "x-anchormailbox": f"PUID:{os.getenv('OWA_PUID')}",
            "x-js-experiment": "5",
            "x-message-count": "25",
            "x-owa-bootflights": "auth-cacheTokenForMetaOsHub,auth-useAuthTokenClaimsForMetaOsHub,auth-codeChallenge,auth-msaljs-eventify,cal-store-NavBarData,auth-msaljs-newsletters,auth-msaljs-meet,auth-msaljs-places,auth-msaljs-bookings,auth-msaljs-findtime,auth-msaljs-landingpage,auth-msaljs-business,auth-msaljs-consumer,pe1416235c1:145887,pe1444542c1:157348,pe1445601c1:157790,shellmultiorigin:394927,cal-perf-useassumeoffset:522141,cal-perf-eventsinofflinestartupdata:549998,acct-add-account-e1-improvement:515788,auth-cachetokenformetaoshub:579266,cal-perf-eventsinstartupdatabyviewtype:542451,disableconcurrency_cf:777754,auth-msaljs-business:757509,workerasyncload:634921,auth-msaljs-newsletters:750219,auth-msaljs-bookings:795667,auth-msaljs-places:657523,cal-reload-pause:774835,auth-useauthtokenclaimsformetaoshub:653460,acctstartdataowav2:700801,auth-msaljs-places-sessionstorage:656411,cal-store-navbardata:683809,acctpersistentidindexerv2:711988,fwk-enforce-trusted-types:864189,fwk-decode-checkexplictlogon:758185,auth-disableanonymoustokenforheadercf:889193,auth-msaljs-meet:750972,auth-msaljs-safari:770018,msplaces-hosted-localsessiondata-v2:776784,auth-codechallenge:768093,msplaces-app-boot-sequence:772374,fwk-nonbootconfig-gql:780302,fwk-enable-default-trusted-types-policy:815211,fwk-enable-redirect-new-olkerror-page:778384,auth-msaljs-hostedplaces:860810,fwk-enable-owa-loop-trusted-types-policy:785464,fwk-init-acc-locstore:828561,fwk-createstore-deffered:823297,auth-msaljs-domainhint:829153,auth-tokencache-expirydate:822074,fwk-getcopilot-fromstartup:821269,auth-mon-removefallbacktogatfrcf:888280,auth-hosted-removefallbacktogatfr:892172,cal-appcaching-reloadonerror:858027,auth-msaljs-consumer-boot-errorhandler:853742,pr161716910:272413,pr1626504c7:272416,pr153989220:305744,pr1634711c7:1000610,auth-enable-prod-icloud-client-id,nh-enableDiagnosticsInteropRefactor,flightship",
            "x-owa-correlationid": os.getenv('OWA_INIT_CORRELATIONID'),
            "x-owa-hosted-ux": "false",
            "x-owa-sessionid": os.getenv('OWA_SESSIONID'),
            "x-req-source": "Mail",
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-origin",
            "authorization": f"Bearer {os.getenv('OWA_BEARER')}",
            "Priority": "u=4"
        },
        "method": "POST",
        "mode": "cors"
    })


def get_emails():
    print('Retrieving emails...')
    conversations = get_emails_raw()['Body']['Conversations']
    with concurrent.futures.ThreadPoolExecutor(max_workers=6) as ex:  # 6 = emulate browsers
        conv_body_pairs = [*ex.map(handle_single_conversation, conversations)]
    for conversation, details in conv_body_pairs:
        # can't do it on different thread due to pickle issues??
        conversation['DETAILS'] = details
    return [get_html_from_email(c) for c in conversations]


def get_html_from_email(conversation: dict[str, Any]):
    # TODO: handle multiple messages in one thread
    nodes = conversation["DETAILS"]["ConversationNodes"]
    if len(nodes) != 1:
        print(f'{len(nodes)=}: {nodes}', file=sys.stderr)
    items = nodes[0]["Items"]
    if len(items) != 1:
        print(f'{len(items)=}: {items}', file=sys.stderr)
    body_obj = items[0]["UniqueBody"]
    if body_obj["BodyType"] != "HTML":
        print(f'BodyType={body_obj["BodyType"]}: {items[0]}', file=sys.stderr)
    return body_obj["Value"]


def get_emails_raw():
    resp_obj = fetch(  # Mainly copy-paste from devtools
        "https://outlook.office.com/owa/service.svc?action=FindConversation&app=Mail&n=5", {
            "credentials": "include",
            "headers": {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:143.0) Gecko/20100101 Firefox/143.0",
                "Accept": "*/*",
                "Accept-Language": "en-GB,en;q=0.5",
                "action": "FindConversation",
                # TODO: generate this? finding client_id from the OWA using devtools hacks
                "authorization": f"Bearer {os.getenv('OWA_BEARER')}",
                "content-type": "application/json; charset=utf-8",
                "ms-cv": f"{os.getenv('OWA_LS_MS_CV')}",
                "prefer": "IdType=\"ImmutableId\", exchange.behavior=\"IncludeThirdPartyOnlineMeetingProviders\"",
                "x-anchormailbox": f"PUID:{os.getenv('OWA_PUID')}",
                "x-owa-correlationid": f"{os.getenv('OWA_LS_CORRELATIONID')}",
                "x-owa-hosted-ux": "false",
                # TODO: this works in mysterious ways (JS client generates it
                #  randomly and registers it woth server? we attempt to emulate
                #  this in init_office()
                "x-owa-sessionid": f"{os.getenv('OWA_SESSIONID')}",
                "x-owa-urlpostdata": urllib.parse.quote(json.dumps({
                    "__type": "FindConversationJsonRequest:#Exchange",
                    "Body": {
                        "ConversationShape": {
                            "__type": "ConversationResponseShape:#Exchange",
                            "BaseShape": "IdOnly"
                        },
                        "FocusedViewFilter": -1,
                        "Paging": {
                            "__type": "SeekToConditionPageView:#Exchange",
                            "BasePoint": "Beginning",
                            "Condition": {
                                "__type": "RestrictionType:#Exchange",
                                "Item": {
                                    "__type": "IsEqualTo:#Exchange",
                                    "FieldURIOrConstant": {
                                        "__type": "FieldURIOrConstantType:#Exchange",
                                        "Item": {
                                            "__type": "Constant:#Exchange",
                                            "Value": "AQAAABjHhTkBAAAAHwIg0wAAAAA="
                                        }
                                    },
                                    "Item": {
                                        "__type": "PropertyUri:#Exchange",
                                        "FieldURI": "ConversationInstanceKey"
                                    }
                                }
                            },
                            "MaxEntriesReturned": 50
                        },
                        "ParentFolderId": {
                            "__type": "TargetFolderId:#Exchange",
                            "BaseFolderId": {
                                "__type": "FolderId:#Exchange",
                                "Id": "AQMkADAzODA5NzA3LWI3NTEtNDAxMC04MmMwLWNjADU4MDQ4MjYyODIALgAAA/uV1hy/6XhAtrtIgOrOGv4BAI2thI0NWwZPqja/whIfpMwAAAIBDAAAAA=="
                            }
                        },
                        "ShapeName": "ReactConversationListView",
                        "SortOrder": [
                            {
                                "__type": "SortResults:#Exchange",
                                "Order": "Descending",
                                "Path": {
                                    "__type": "PropertyUri:#Exchange",
                                    "FieldURI": "ConversationLastDeliveryOrRenewTime"
                                }
                            },
                            {
                                "__type": "SortResults:#Exchange",
                                "Order": "Descending",
                                "Path": {
                                    "__type": "PropertyUri:#Exchange",
                                    "FieldURI": "ConversationLastDeliveryTime"
                                }
                            }
                        ],
                        "ViewFilter": "All"
                    },
                    "Header": {
                        "__type": "JsonRequestHeaders:#Exchange",
                        "RequestServerVersion": "V2018_01_08",
                        "TimeZoneContext": {
                            "__type": "TimeZoneContext:#Exchange",
                            "TimeZoneDefinition": {
                                "__type": "TimeZoneDefinitionType:#Exchange",
                                "Id": "GMT Standard Time"
                            }
                        }
                    }
                })
                ),
                "x-req-source": "Mail",
                "Sec-Fetch-Dest": "empty",
                "Sec-Fetch-Mode": "cors",
                "Sec-Fetch-Site": "same-origin",
                "Priority": "u=4"
            },
            "method": "POST",
            "mode": "cors"
        })
    if not resp_obj.ok:
        need_reauth()
    try:
        resp = resp_obj.json()
    except requests.exceptions.JSONDecodeError as e:
        if not any(resp_obj.content):
            need_reauth()  # Empty or all NUL bytes = need re-auth
        print(f'ERR! {e}')
        print(f'ERR! src: {resp_obj.content}')
        raise
    if resp.get('Body').get("ResponseClass") != "Success":
        raise ApiError(resp)
    return resp


def need_reauth() -> NoReturn:
    print('Re-auth is needed!')
    os.kill(os.getpid(), signal.SIGINT)  # errm...
    sys.exit(2)


# dicts aren't hashable (me sad...) so we use this homemade function
@diskcache('.app_cache/email_thread_cache.pkl',
           lambda conversation_id: conversation_id['Id'])
def get_email_thread(conversation_id: dict[str, ...]):
    resp = fetch(  # Once again, mainly copy-paste from devtools
        "https://outlook.office.com/owa/service.svc?action=GetConversationItems&app=Mail&n=22",
        {
            "credentials": "include",
            "headers": {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:143.0) Gecko/20100101 Firefox/143.0",
                "Accept": "*/*",
                "Accept-Language": "en-GB,en;q=0.5",
                "action": "GetConversationItems",
                "authorization": f"Bearer {os.getenv('OWA_BEARER')}",
                "content-type": "application/json; charset=utf-8",
                "ms-cv": os.getenv('OWA_TH_MS_CV'),
                "prefer": "IdType=\"ImmutableId\", exchange.behavior=\"IncludeThirdPartyOnlineMeetingProviders\"",
                "x-anchormailbox": f"PUID:{os.getenv('OWA_PUID')}",
                "x-owa-correlationid": os.getenv('OWA_CORRELATIONID'),
                "x-owa-hosted-ux": "false",
                "x-owa-sessionid": os.getenv('OWA_SESSIONID'),
                "x-req-source": "Mail",
                "Sec-Fetch-Dest": "empty",
                "Sec-Fetch-Mode": "cors",
                "Sec-Fetch-Site": "same-origin",
                "Priority": "u=4"
            },
            "body": json.dumps({
                "__type": "GetConversationItemsJsonRequest:#Exchange",
                "Header": {
                    "__type": "JsonRequestHeaders:#Exchange",
                    "RequestServerVersion": "V2017_08_18",
                    "TimeZoneContext": {
                        "__type": "TimeZoneContext:#Exchange",
                        "TimeZoneDefinition": {
                            "__type": "TimeZoneDefinitionType:#Exchange",
                            "Id": "GMT Standard Time"
                        }
                    }
                },
                "Body": {
                    "__type": "GetConversationItemsRequest:#Exchange",
                    "Conversations": [
                        {
                            "__type": "ConversationRequestType:#Exchange",
                            "ConversationId": conversation_id,
                            "SyncState": ""
                        }
                    ],
                    "ItemShape": {
                        "__type": "ItemResponseShape:#Exchange",
                        "BaseShape": "IdOnly",
                        "AddBlankTargetToLinks": True,
                        "BlockContentFromUnknownSenders": False,
                        "BlockExternalImagesIfSenderUntrusted": True,
                        "ClientSupportsIrm": True,
                        "CssScopeClassName": "rps_efc5",
                        "FilterHtmlContent": True,
                        "FilterInlineSafetyTips": True,
                        "InlineImageCustomDataTemplate": "{id}",
                        "InlineImageUrlTemplate": "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAEALAAAAAABAAEAAAIBTAA7",
                        "MaximumBodySize": 2097152,
                        "MaximumRecipientsToReturn": 20,
                        "ImageProxyCapability": "OwaAndConnectorsProxy",
                        "AdditionalProperties": [
                            {
                                "__type": "PropertyUri:#Exchange",
                                "FieldURI": "CanDelete"
                            },
                            {
                                "__type": "ExtendedPropertyUri:#Exchange",
                                "PropertySetId": "00062008-0000-0000-C000-000000000046",
                                "PropertyName": "ExplicitMessageCard",
                                "PropertyType": "String"
                            },
                            {
                                "__type": "ExtendedPropertyUri:#Exchange",
                                "PropertySetId": "E550B918-9859-47B9-8095-97E4E72F1926",
                                "PropertyName": "IOpenTypedFacet.Com.Microsoft.Graph.MessageCard",
                                "PropertyType": "String"
                            },
                            {
                                "__type": "ExtendedPropertyUri:#Exchange",
                                "PropertyName": "DrawingCanvasElements",
                                "DistinguishedPropertySetId": "Common",
                                "PropertyType": "String"
                            },
                            {
                                "__type": "ExtendedPropertyUri:#Exchange",
                                "PropertyName": "TemplateName",
                                "DistinguishedPropertySetId": "Common",
                                "PropertyType": "String"
                            },
                            {
                                "__type": "ExtendedPropertyUri:#Exchange",
                                "PropertyName": "OpenedProperty",
                                "DistinguishedPropertySetId": "Common",
                                "PropertyType": "String"
                            },
                            {
                                "__type": "PropertyUri:#Exchange",
                                "FieldURI": "OwnerReactionType"
                            },
                            {
                                "__type": "PropertyUri:#Exchange",
                                "FieldURI": "Reactions"
                            },
                            {
                                "__type": "ExtendedPropertyUri:#Exchange",
                                "PropertyName": "NetworkMessageId",
                                "DistinguishedPropertySetId": "Common",
                                "PropertyType": "CLSID"
                            },
                            {
                                "__type": "ExtendedPropertyUri:#Exchange",
                                "PropertyName": "X-MS-Exchange-Organization-ATPSafeLinks-MsgData",
                                "DistinguishedPropertySetId": "Common",
                                "PropertyType": "String"
                            },
                            {
                                "__type": "ExtendedPropertyUri:#Exchange",
                                "PropertyName": "X-MS-Reactions",
                                "DistinguishedPropertySetId": "InternetHeaders",
                                "PropertyType": "String"
                            },
                            {
                                "__type": "PropertyUri:#Exchange",
                                "FieldURI": "HasProcessedSharepointLink"
                            },
                            {
                                "__type": "PropertyUri:#Exchange",
                                "FieldURI": "CopilotInboxScoreReason"
                            },
                            {
                                "__type": "PropertyUri:#Exchange",
                                "FieldURI": "CopilotInboxHeadline"
                            },
                            {
                                "__type": "PropertyUri:#Exchange",
                                "FieldURI": "IsSendIndividually"
                            }
                        ],
                        "InlineImageUrlOnLoadTemplate": "",
                        "ExcludeBindForInlineAttachments": True,
                        "CalculateOnlyFirstBody": True,
                        "BodyShape": "UniqueFragment"
                    },
                    "ShapeName": "ItemPart",
                    "SortOrder": "DateOrderAscending",
                    "MaxItemsToReturn": 20,
                    "Action": "ReturnRootNode",
                    "FoldersToIgnore": [
                        {
                            "__type": "FolderId:#Exchange",
                            "Id": "AQMkADAzODA5NzA3LWI3NTEtNDAxMC04MmMwLWNjADU4MDQ4MjYyODIALgAAA/uV1hy/6XhAtrtIgOrOGv4BAI2thI0NWwZPqja/whIfpMwAAAIBEAAAAA=="
                        },
                        {
                            "__type": "FolderId:#Exchange",
                            "Id": "AQMkADAzODA5NzA3LWI3NTEtNDAxMC04MmMwLWNjADU4MDQ4MjYyODIALgAAA/uV1hy/6XhAtrtIgOrOGv4BAI2thI0NWwZPqja/whIfpMwAAAIBCwAAAA=="
                        },
                        {
                            "__type": "FolderId:#Exchange",
                            "Id": "AAMkADAzODA5NzA3LWI3NTEtNDAxMC04MmMwLWNjNTgwNDgyNjI4MgAuAAAAAAD7ldYcv+l4QLa7SIDqzhr+AQCNrYSNDVsGT6o2v8ISH6TMAAAUnguRAAA="
                        }
                    ],
                    "ReturnSubmittedItems": True,
                    "ReturnDeletedItems": True
                }
            }),
            "method": "POST",
            "mode": "cors"
        })
    if not resp.ok:
        raise ApiError(resp)
    try:
        return resp.json()["Body"]["ResponseMessages"]["Items"][0]["Conversation"]
    except KeyError:
        writefile('err.json', resp.content.decode('utf8'))
        raise


HAS_INITIAL_CHAT_MSG = False


# We can't cache this as we need the chat context
# @diskcache('.app_cache/summ_pdfs.pkl', lambda mails: tuple(mails))
def tasklist_from_pdfs(pdfs: list[bytes]):
    global HAS_INITIAL_CHAT_MSG, chat  # not good practise but it's late.
    print('Summarising tasks...')
    chat = genai_client.chats.create(model='gemini-2.5-flash')  # meow!
    res = chat.send_message(
        [
            'Summarise the tasks the the user needs to perform based on '
            'these emails. Put them in order of priority (with the most '
            'important/urgent one at the front of the array). The output '
            'should be JSON conforming to the schema. Ignore any events '
            f'in the past (today is {datetime.datetime.now().strftime("%d/%m/%Y")}'
            f', time is {datetime.datetime.now().strftime("%H:%M")}).'
            f'Unless otherwise specified, prioritise academic tasks, such '
            f'as lectures and supervisions.',
            *[genai.types.Part.from_bytes(data=pdf, mime_type='application/pdf')
              for pdf in pdfs],

        ],
        config=genai.types.GenerateContentConfig(response_schema={
            "type": "ARRAY",
            "items": {
                "type": "STRING"
            }
        }, response_mime_type='application/json')
    )
    HAS_INITIAL_CHAT_MSG = True
    return res.parsed


def update_tasklist(info: str):
    print('Updating list...')
    res = chat.send_message(
        f"The user has provided new information: {info}. "
        f"Please reorder the list based on this new information, "
        f"also removing any items that are completely irrelevant"
        f" to the user (for example, if the user doesn't care"
        f" about music, don't include anything related to music,"
        f" unless explicitly unequivocally compulsory). If not action needs"
        f" to be taken, you should omit that item completely.",
        config=genai.types.GenerateContentConfig(
            response_schema={
                "type": "ARRAY",
                "items": {
                    "type": "STRING"
                }
            }, response_mime_type='application/json'
        )
    )
    return res.parsed


def handle_single_conversation(conversation: dict):
    return conversation, get_email_thread(conversation_id=conversation['ConversationId'])


@app.route('/')
def index():
    if info := request.args.get('user_info'):
        if HAS_INITIAL_CHAT_MSG:
            tasklist = update_tasklist(info)
            return render_template('index.html', tasklist=tasklist)
        else:
            return redirect('/')  # It doesn't have the emails yet
    print('Initialising...')
    init_office()
    emails = get_emails()
    print('Converting documents to PDF...')
    pdfs = [html_to_pdf(mail) for mail in emails]
    tasklist = tasklist_from_pdfs(pdfs)
    return render_template('index.html', tasklist=tasklist)
