#! python3
# Get input from excel file -> Find email -> Directly update excel file 

import sys
import requests
import re
import openpyxl
import os
from bs4 import BeautifulSoup
from googlesearch import search

# TODO: handle Email Address Obfuscation
# TODO: bypass Incapsula
# TODO: handle "farinaz (at) ucsd (dot) edu"/"girard at lirmm dot fr"/"gunnar [dot] liden [at] chemeng [dot] lth [dot] se" email


os.chdir(r"C:/Users/HP/MyPythonScripts/vin_future")

def is_valid_email(email):
    invalid = [
        'admin', 'admissions', 'app', 'ask', 'biochem', 'business', "biosci",
        'career', 'college', 'communications', 'contact', 'contact', 'copy', "council",
        'director', 'engineer', 'enquiries', 'experts', "feedback", "financial", 'help', 
        'idm', 'idm', 'info', 'inquiry', 'international', "institution", 'job', 'kontakt', 
        'lab', 'life', 'news', 'medcentral', "medicine", 'office', "order",
        'phd', 'president', 'press', 'profile', 'profiles', 'program', "physic", 
        'registrator', 'reply', 'research', 'researcher', 
        'sales', 'service', 'staff', 'support', "student", "science",
        'team', 'web', "xxx"]   
    first, _ = email.split('@')
    for invalid_word in invalid:
        if invalid_word in first.lower():
            return False
    return True


def extract_mailto(soup):
    emails = []
    mailtos = soup.select('a[href^="mailto"]')
    if len(mailtos) == 0:
        mailtos = soup.select('a[href^="Mailto"]')
    for i in mailtos:
        href = i['href']
        try:
            _, mail = href.split(':')
        except ValueError as e:
            print(e)
            continue

        if not is_valid_email(mail):
            continue

        emails.append(mail)
    print(emails)
    return emails



def extract_mail_reg(soup):
    text = soup.get_text()
    print(text)
    """
    email_regex = re.compile(r'''
    ^([a-z0-9]+(?:[.-]?[a-z0-9]+)*    # username
    (\[|\s|\()*at+(\]|\s|\))* # handle @ = at
    @*
    [a-z0-9]+(?:[.-]?[a-z0-9]+)*    # domain
    (\[|\s|\()*[dot]*(\]|\s|\))*    # handle . = dot
    \.*[a-z]{2,7})$ # domain 
    ''', re.VERBOSE | re.IGNORECASE | re.MULTILINE)
    match = email_regex.findall(text)
    """

    # flatten nested findall
    #flatten(match)

    '''
    #match = re.findall(r'[\w\.-]+@[\w\.-]+', text)
    email_regex_1 = re.compile(r"""
    ^([a-z0-9]+(?:[.-]?[a-z0-9]+)*
    ( \[ | \( *[at]*( \] | \) )* # handle @ = at   
    [a-z0-9]+(?:[.-]?[a-z0-9]+)* 
    \.[a-z]{2,7})$
    """, re.VERBOSE | re.IGNORECASE | re.MULTILINE)

    match = email_regex_1.findall(text)
    '''

    
    email_regex_2 = re.compile(r'''
    ^([a-z0-9]+(?:[.-]?[a-z0-9]+)*
    @
    [a-z0-9]+(?:[.-]?[a-z0-9]+)*
    \.[a-z]{2,7})$
    ''', re.VERBOSE | re.IGNORECASE)

    match = email_regex_2.findall(text)
    
    
    if len(match) == 0:
        match = re.findall(r'[\w\.-]+@[\w\.-]+', text)
    match = [e for e in match if is_valid_email(e)]

    print(match)
    return match


urls = ["https://www.lunduniversity.lu.se/lucat/user/b0db29b738c0d2259422bca76ad4ed68"]
def find_email(urls):
    possible_emails = []
    # scrape for the first 2 websites / 2 seconds pause
    for url in urls:
        print(url)

        if "researchgate" in url:
            print("research gate")
            continue
        # access to the url, move to the next web if any errors
        try:
            page = requests.get(url, timeout=10)
        except Exception as e:
            print(e)
            continue

        # exclude pdf file
        content_type = page.headers.get("content-type")
        if "application/pdf" in content_type:
            print("pdf file")
            continue

        soup = BeautifulSoup(page.content, "html.parser")
        print(soup)

        # find email by mailto html
        try:
            print("mailto")
            emails = extract_mailto(soup)
            possible_emails.extend(emails)
            if len(possible_emails) == 0:
                raise Exception("Not Found")
        except:
            print("re")
            # find email using regular expression
            match = extract_mail_reg(soup)
            possible_emails.extend(match)

    return possible_emails

find_email(urls)

