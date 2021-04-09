import sys
import requests
import re
from requests.auth import HTTPBasicAuth
from datetime import datetime, timedelta
from utils import * 
from sheets import *
import pickle

DEBUG = False 
REPORT_URL = "https://app.cybersmart.co.uk/1070ca00-6f37-4f37-910e-099381ed7bf7/software-report/"
BASE_URL = "https://app.cybersmart.co.uk"
def _get_headers():
    headers = {
        "accept-encoding": "gzip, deflate",
        "accept-language": "en-GB,en-US;q=0.9,en;q=0.8",
        "cookie": "isiframeenabled=true; isiframeenabled=true; __cfduid=d6db32edc99b860e5f23b8ce44c7965021616670738; _ga=GA1.3.1856284293.1616670740; _gid=GA1.3.1334888485.1616670740; ajs_anonymous_id=%2297f302b3-43f2-473f-8a89-709d98f10558%22; _fbp=fb.2.1616670742601.966516833; _gcl_aw=GCL.1616671070.CjwKCAjw6fCCBhBNEiwAem5SO3H3nuluuDFT-GPTDwKwyAyanfIFMKCs1fHjPzQ9KPp_TIjhL8kHTBoCbNoQAvD_BwE; gclid=undefined; cybersmartlimited-_zldp=knG7VsVQcTgCZ9CfO1%2F%2F388xK8V8wZKizoNFqyogeVbCgzXS3mDkseybyThG5U3elLnzf3o1QSQ%3D; cybersmartlimited-_zldt=504da75a-9ed4-47ae-9e62-93b46fd77b16-0; sessionid=9o803avhh7ch73hj6fvgnce8hb4crntj; _iub_cs-8123391=%7B%22timestamp%22%3A%222021-03-25T13%3A56%3A27.929Z%22%2C%22version%22%3A%221.28.1%22%2C%22purposes%22%3A%7B%221%22%3Atrue%2C%222%22%3Atrue%2C%223%22%3Atrue%2C%224%22%3Atrue%2C%225%22%3Atrue%7D%2C%22id%22%3A8123391%7D; euconsent-v2=CPDnGDePDneF4B7D6BENBSCsAP_AAH_AAAAAHnNf_X__b39j-_59_9t0eY1f9_7_v-0zjhfds-8N2f_X_L8X42M7vF36pq4KuR4Eu3LBIQFlHOHUTUmw6okVrTPsak2Mr7NKJ7LEinMbe2dYGHtfn91TuZKYr_78_9_z__-__v__79f_r-3_3_vp9X---_e_V399xLv9cDygCTDUvgAsxLHBkmjSqFECEK4kOgFABRQjC0TWEDK4KdlcBHqCBgAgNQEYEQIMQUYsAgAAAgCSiICQA8EAiAIgEAAIAVICEABEwCCwAsDAIABQDQsQIoAhAkIMjgqOUwICJFooJ5KwBKLvY0whDKLACgUf0VGAiUIIFgAA; _gac_UA-90566725-1=1.1616680614.CjwKCAjw6fCCBhBNEiwAem5SO4f6LiPEwf5XVbZqElJHvdF_zEhjGorC46V6B0kz_dkVZR3PkDb4lxoCTxMQAvD_BwE; __stripe_mid=fff6cbbb-99b6-4be5-b768-3d6e85aa17b476ef5d; _iub_cs-8123391-granular=%7B%22gac%22%3A%22MX4mAQMBAgEIAQUBBAEDAQwBBQEDAQ4BCAEEAQEBBgEDAgYCAgEBAQkBAgEEAQMBFAEDAQUBCAEGAQkBAQEIAQEBCwEFAQYBBQENAQQBEwEFAQQCAgIKARwBAwENAQMBBAECAQkBBQEBAQgBBQEFAQMBBAECAgMBHAEDAQQBAgIFAQEBAQEQAhABCQEIAgcBBQEBAQcBAgEDAcKNAQMBBwEiAQYBDgINAQICBwEJAQ0BCgECAQYBGAEEAREBCAEGASgBAQEDAREBFAECAQMBAQIFAQUBBAEBAQ0BEQEGAQIBAQEBAQcBEwEHAQcBBQECAQkBAwETAQEBAwEIAQICBQEDAQoCAgEVAQ8BAQEFAQcBAQEDAQoBBQEEARABDwEKAQcCCQEbAQcBAwICARQBAQEGAQUCCAEdARABAwEIAQ4BBwEGAQIBBQEFAQYBAQEDAgYBCwEMARUBDAEBAQsBAQEDAQUCAwEOAQEBAwEIAQMBBAICAQYBDAEEAQIBCwEDAQwBAwENAQcBAQEOAQEBBAEEAgEBAQIBAQ0BBgEDAQcBAQEIAQkBEQELAQwBAQERAwIDCAEYAQMBEwEHAQMBBAEBAgQBAwEHAQMBAQEBAQEBDQEBAQwBAwEBAQUBCAEFAQIBAwECAQQBAQECAQUBCQEKAQEBAwECAQ8BAgEKAQICAQECAQgBEgEKAQ4BAgEJAQYBBQEDAQIBAwEIAQIBAgECBAUBCgECAwYBAwEFAgkBBAEBAQUBAgEBAQEBAwECAQEBAQEDAQIBAQEMAQYBCwEBAgUBAwEEAQMBAgEBAQEBAwICBAEBCAIFAQgCBAEBAgYBAQEHAQoCAgMBAgIBAQEFAgQBBQIEAQICAgMBAQEBBgEGAgMCAQEFAgEDAgMDAwECBwICAQMBAwECAQECAgIDAQIBCAEFAg4BCQEbAgEDCwECAQMCBQECAQMBBgICAwICBAECAgIBAQEBAwMBAQIBBAEBAQEDAQECAQEBAQEBAwQBDAEBAQMBAgECBgIBBAEFAQMCAQEDAQEFBAQBAQIFAQQGAQEBAQIDAwMBAQEDAwEBAwEBAgEBAgEBAgUDBAQCCAECAgQDAgEDAgEBAgIDAQECAQEBAQEHCgICAwEBAQIDAQEBAgEBAgMGAgEBAwEBBAICAQUBAwMCAgMBAgIDAQQCBgEBBAEEAgIBAgsDAgIBAQEBAgICAwEBAQEBAQED%22%7D; ZLD6d7ae9a9b775a2cda0ccf872e6283ff1299efa3a6cd5f15ctabowner=undefined; _gat_gtag_UA_90566725_1=1; _uetsid=f81367b08d5a11eb9e4bf1de025a9c9f; _uetvid=f81384f08d5a11ebac6281dff5ebff85",
        "referer": "https://app.cybersmart.co.uk/1070ca00-6f37-4f37-910e-099381ed7bf7/software-report/",
        "sec-ch-ua": '"Google Chrome";v="89", "Chromium";v="89", ";Not A Brand";v="99"',
        "sec-ch-ua-mobile": "?0",
        "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "same-origin",
        "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 11_2_0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.90 Safari/537.36",
        "x-requested-with": "XMLHttpRequest"
    }
    return headers

def _clean(l):
    return l.replace("<td>","").replace("</td>","").strip()

def _get_devices(l):
    re_href = re.compile('data-href=(.+)>')
    re_href = re.compile('.*data-href="(.*)">')
    m = re_href.match(l)
    if not m or len(m.groups())==0:
        print("No match")
        sys.exit()
        return None
    href = m.groups()[0]
    url = BASE_URL+href
    data = _make_get(url, json=True)
    return data
    

def _make_get(req_url, qs=None, json=False):
    resp = requests.get(req_url, headers=_get_headers(), params=qs)
    resp.raise_for_status()
    if resp.status_code==204:
        print("Nothing to return")
        return None
    else: 
        if json:
            return resp.json()
        else:
            return resp.text

def _get_rows(): 
    cache_key = "sec-apps-rows"
    row_list = get_cache(cache_key)
    if not row_list:
        row_list = []
        for i in range(1,100):
            
            page_qs = {
                "page": i
            }
            txt = _make_get(REPORT_URL, page_qs)    
            print(f"Doing page {i} length {len(txt)}")
            if len(txt)<10:
                break
            row = {}
            for l in txt.split("\n"):
                if l.strip()=="":
                    continue
                if l=="<tr >":
                    if row:
                        row_list.append(row)
                    row = {}
                    col = 0 
                    vuln = False
                else:
                    col += 1
                    if vuln:
                        print(f"{col} - {l}")
                    if col==2:
                        row['vuln'] = _clean(l)
                        if row['vuln']!="No known vulnerabilities":
                            vuln = True
                    elif (not vuln and col==4) or (vuln and col==5):
                        row['sw'] = _clean(l)
                    elif (not vuln and col==5) or (vuln and col==6):
                        row['ver'] = _clean(l)
                    elif (not vuln and col==6) or (vuln and col==7):
                        row['vendor'] = _clean(l)
                    elif (not vuln and col==7) or (vuln and col==8):
                        devices = _get_devices(_clean(l))
                        row['swref'] = _clean(l)
                        row['hosts'] = devices['hosts']
        set_cache(cache_key, row_list)          
    return row_list

row_list = _get_rows()
app_sheet = SheetService("1P3T41uVnwr3UZZPXFmNeXJL8IaPQb0yGU2ocok04qvc")      
#print_json(row_list) 
print(len(row_list))

summary_list = []
sd = app_sheet.get("Apps!A1:G1000")
for sr in sd:
    #print(sr)
    if sr[0]=="Vendor": 
        continue
    if len(sr)>=4 and sr[3]=="x":
        report = True
    else:
        report = False
    summary_list.append({
        "vendor": sr[0],
        "sw": sr[1],
        "versions": 0,
        "report": report
    })
    


for r in row_list:
    matched = False
    for s in summary_list:
        if s['vendor']==r['vendor'] and s['sw']==r['sw']:
            matched = True
            s['versions'] += 1
            break
    if not matched:
        print(f"New software {r['vendor']} {r['sw']}")
        summary_list.append({
            "vendor": r['vendor'],
            "sw": r['sw'],
            "versions": 1
        })

summary_header_map = {
    "vendor": "Vendor",
    "sw": "Software",
    "versions": "Versions"
}
app_sheet.build_range("Apps",summary_header_map, summary_list)

rows_header_map = {
    "vendor": "Vendor",
    "sw": "Software",
    "ver": "Version",
    "vuln": "Vulnerability"
}
app_sheet.build_range("Full", rows_header_map, row_list)

devices = {}
hosts = []
host_ids = []
host_map = {
   "DESKTOP-8218CKI": "Alex",
   "Cristis-MacBook-Pro.local": "Cristian",
   "Ruairidhs-MacBook-Pro.local": "Ruairidh",
   "MacBook-Pro.local": "Sam",
   "Sleipnir.localdomain": "Simon",
   "Kuhelies-MacBook-Pro.local": "Kuhele",
   "Thomass-MacBook.local": "Shrive",
   "DESKTOP-09K8LG6": "Charlotte",
   "Fabios-MacBook-Pro.local": "Fabio",
   "Jamies-Air": "Jamie",
   "toms-mbp.lan": "Carter",
   "Elenas-MBP": "Elena",
   "Lees-Air": "Lee"
}
for r in row_list:
    for h in r['hosts']:
        if h['hostname'] not in hosts:
            hosts.append(h['hostname'])
summary_map = {}
for s in summary_list:        
    summary_map[f"{s['vendor']}-{s['sw']}"] = s
user_data = {}
user_data_all = {}
for r in row_list:
    key = f"{r['vendor']}-{r['sw']}"
    if key not in summary_map:
        print(f"MIssing {key}")
        sys.exit()
    #print_json(summary_map[key])
    if summary_map[key]['report']:
        if key not in user_data:
            user_data[key] = {
                "vendor": r['vendor'],
                "sw": r['sw']
            }
            for uk in host_map:
                user_data[key][host_map[uk]] = ""
        for h in r['hosts']:
            if h['hostname'] not in host_map:
                print("Missing hos")
                sys.exit()
            usr = host_map[h['hostname']]
            user_data[key][usr] = r['ver']
    

    if key not in user_data_all:
        user_data_all[key] = {
            "vendor": r['vendor'],
            "sw": r['sw']
        }
        for uk in host_map:
            user_data_all[key][host_map[uk]] = ""
    for h in r['hosts']:
        if h['hostname'] not in host_map:
            print("Missing hos")
            sys.exit()
        usr = host_map[h['hostname']]
        user_data_all[key][usr] = r['ver']
#print_json(user_data)

user_header_map = {
    "vendor": "Vendor",
    "sw": "Software",
   "Alex": "Alex",
   "Cristian": "Cristian",
   "Ruairidh": "Ruairidh",
   "Sam": "Sam",
   "Simon": "Simon",
   "Kuhele": "Kuhele",
   "Shrive": "Shrive",
   "Charlotte": "Charlotte",
   "Fabio": "Fabio",
   "Jamie": "Jamie",
   "Carter": "Carter",
   "Elena": "Elena",
   "Lee": "Lee"
}
user_data_list = sorted(user_data.values(), key=lambda k: (k['vendor'], k['sw']))
app_sheet.build_range("Users", user_header_map, user_data_list)

user_data_all_list = sorted(user_data_all.values(), key=lambda k: (k['vendor'], k['sw']))
app_sheet.build_range("Users All", user_header_map, user_data_all_list)

