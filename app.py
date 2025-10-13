"""App to sort emails by """
import concurrent.futures
import datetime
import functools
import json
import os
import pickle
import signal
import sys
import threading
import time
import urllib.parse
from collections.abc import Callable, Hashable
from datetime import timedelta
from typing import TypeVar, ParamSpec, Any, NoReturn

import requests
from dotenv import load_dotenv
from flask import Flask, Response, render_template, request, redirect
from google import genai
import pdfkit


load_dotenv()
app = Flask(__name__)
genai_client = genai.Client()  # TODO: if lite no enuf 4 dis, change here vvv
chat = genai_client.chats.create(model='gemini-2.5-flash')  # meow!

# I can't believe ParamSpec was added as early as 3.10 - it feels like such a 3.13 thing
P = ParamSpec('P')
R = TypeVar('R')
H = TypeVar('H', bound=Hashable)


def _default_cache_key(*args, **kwargs):
    return args, frozenset(kwargs)


def diskcache(filename: str = None, key_fn: Callable[P, H] = _default_cache_key,
              lifetime: datetime.timedelta | float | None = datetime.timedelta(minutes=10)):
    # NOTE: not thread-safe in the slightest! Also, the performance is A LOT
    # worse than functools.cache but this one should be used for very expensive
    # operations (e.g. requesting data from the web)
    lifetime_sec = (
        lifetime.total_seconds() if isinstance(lifetime, datetime.timedelta)
        else datetime.timedelta(days=1e15) if lifetime is None else lifetime)

    def decor(fn: Callable[P, R]) -> Callable[P, R]:
        try:
            if filename is not None:
                with open(filename, 'rb') as fr:
                    cache: dict[H, tuple[float, R]] = pickle.load(fr)
            else:
                cache = {}
        except FileNotFoundError:  # create cache file after first entry added
            cache = {}
        cache_lock = threading.RLock()

        @functools.wraps(fn)
        def new_fn(*args, **kwargs):
            with cache_lock:
                key = key_fn(*args, **kwargs)
                try:  # Could also use contextlib.suppress here but this is clearer
                    birth, value = cache[key]
                    if birth + lifetime_sec > time.time():
                        print(f'(Cache hit for {filename})')
                        return value
                    print(f'(Cache expired for {filename})')
                    del cache[key]  # don't leak even if error below
                except KeyError:
                    print(f'(Cache miss for {filename})')
            value = fn(*args, **kwargs)
            birth = time.time()
            with cache_lock:
                cache[key] = birth, value
                if filename:
                    with open(filename, 'wb') as fw:
                        # Pycharm still doesn't understand Protocol after 6 years!
                        # noinspection PyTypeChecker
                        pickle.dump(cache, fw)
            return value
        return new_fn
    return decor


class ApiError(Exception):
    def __init__(self, o: object):
        super().__init__(json.dumps(o, indent=2))


@diskcache('.app_cache/topdf_cache.pkl', lifetime=timedelta(days=30))
def topdf(html: str) -> bytes:
    return pdfkit.from_string(html, configuration=pdfkit.configuration(
        wkhtmltopdf='C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe'))


HAS_INITIAL_CHAT_MSG = False


# We can't cache this as we need the chat context
# @diskcache('.app_cache/summ_pdfs.pkl', lambda mails: tuple(mails))
def summ_pdfs(pdfs: list[bytes]):
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


def need_reauth() -> NoReturn:
    print('Re-auth is needed!')
    os.kill(os.getpid(), signal.SIGINT)  # errm...
    sys.exit(2)


def get_email():
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
    if not (200 <= resp_obj.status_code < 300):
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


# dicts aren't hashable (me sad...) so we use this homemade function
@diskcache('.app_cache/email_thread_cache.pkl', lambda convid: convid['Id'])
def get_email_thread(convid: dict[str, ...]):
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
                            "ConversationId": convid,
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
    if resp.status_code > 299:
        raise ApiError(resp)
    try:
        return resp.json()["Body"]["ResponseMessages"]["Items"][0]["Conversation"]
    except KeyError:
        writefile('err.json', resp.content.decode('utf8'))
        raise


def fetch(url, conf: dict[str, ...]) -> requests.Response:
    return requests.request(
        conf['method'], url, params=conf.get('params'),
        headers=conf.get('headers'), data=conf.get('body'))


def get_html_from_email(conv: dict[str, Any]):
    # TODO: handle multiple messages in one thread
    nodes = conv["DETAILS"]["ConversationNodes"]
    if len(nodes) != 1:
        print(f'{len(nodes)=}: {nodes}', file=sys.stderr)
    items = nodes[0]["Items"]
    if len(items) != 1:
        print(f'{len(items)=}: {items}', file=sys.stderr)
    body_obj = items[0]["UniqueBody"]
    if body_obj["BodyType"] != "HTML":
        print(f'BodyType={body_obj["BodyType"]}: {items[0]}', file=sys.stderr)
    return body_obj["Value"]


def writefile(fnm: str, data: str):
    with open(fnm, 'w', encoding='utf8') as f:
        f.write(data)


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


def update_list(info: str):
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


def handle_single_conv(conv: dict):
    return conv, get_email_thread(convid=conv['ConversationId'])


def get_emails():
    print('Retrieving emails...')
    convs = get_email()['Body']['Conversations']
    with concurrent.futures.ThreadPoolExecutor(max_workers=6) as ex:  # 6 = emulate browsers
        ops = [*ex.map(handle_single_conv, convs)]
    for dest, val in ops:
        dest['DETAILS'] = val  # can't do it on different thread due to pickle issues??
    return [conv["DETAILS"]["ConversationNodes"][0]["Items"][0]
            ["UniqueBody"]["Value"] for conv in convs]


@app.route('/')
def index():
    if info := request.args.get('user_info'):
        if HAS_INITIAL_CHAT_MSG:
            summs = update_list(info)
            return render_template('index.html', summs=summs)
        else:
            return redirect('/')  # It doesn't have the emails yet
    print('Initialising...')
    init_office()
    emails = get_emails()
    print('Converting documents to PDF...')
    pdfs = [topdf(mail) for mail in emails]
    summs = summ_pdfs(pdfs)
    return render_template('index.html', summs=summs)
