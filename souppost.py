#!/usr/bin/python
# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-

import requests
import random

# url = "http://47.96.83.53:9090/c2in/services/ctms"
url = "http://61.130.250.36:8090/c2in/services/ctms"

headers = {'Content-Type': 'text/xml; charset=utf-8',
           'Accept': 'application/soap+xml, application/dime, multipart/related, text/*',
           'User-Agent': 'Axis/1.4',
           'Cache-Control': 'no-cache',
           'Pragma': 'no-cache',
           'SOAPAction': ''
           }

body = """<soapenv:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:iptv="iptv">
   <soapenv:Header/>
   <soapenv:Body>
      <iptv:ExecCmd soapenv:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
         <CSPID xsi:type="soapenc:string" xmlns:soapenc="http://schemas.xmlsoap.org/soap/encoding/">IPTVZJ</CSPID>
         <LSPID xsi:type="soapenc:string" xmlns:soapenc="http://schemas.xmlsoap.org/soap/encoding/">XWZXIPTVZJ</LSPID>
         <CorrelateID xsi:type="soapenc:string" xmlns:soapenc="http://schemas.xmlsoap.org/soap/encoding/">${CorrelateID}</CorrelateID>
         <CmdFileURL xsi:type="soapenc:string" xmlns:soapenc="http://schemas.xmlsoap.org/soap/encoding/">ftp://c2ftp:c2ftp123@172.24.5.133/${CmdFileURL}</CmdFileURL>
      </iptv:ExecCmd>
   </soapenv:Body>
</soapenv:Envelope>"""


def send_soap(CorrelateID, CmdFileURL):
    post_body = body.replace("${CorrelateID}", CorrelateID).replace("${CmdFileURL}", CmdFileURL)
    print(post_body)
    response = requests.post(url, data=post_body, headers=headers)
    return response.content


if __name__ == '__main__':
    # sequnceid = random.randint(1, 100000000)
    # resp = send_soap(str(sequnceid), '690.0.xml')
    resp = send_soap(str(1814), '1814.0.xml')
    print(resp)