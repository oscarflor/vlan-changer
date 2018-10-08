# vlan-changer
Periodically polls a Google Sheet that stores results of Google Form submissions. If new entries are found, changes are made accordingly.

When VLANs need to be changed, L1 techs submit a form like this https://docs.google.com/forms/d/e/1FAIpQLSfzMkmEeADoViCdI-nbx51dCAs-wvf5KEDOvnxHZ8oOxdwlng/viewform?usp=sf_link

Responses get sent to a google sheet that looks like this https://docs.google.com/spreadsheets/d/1-d15BZDD0NhyisdFBpF38robQLzqwilgcAJB3mdRrac/edit?usp=sharing

Network admins make the necessary changes based on the info in the sheet.

This script automates this process by polling the sheet for new entries. If one or more are found, ssh connections are established using the Netmiko library and changes are made. Emails are sent to the submitters for successful changes. Net admins receive emails for failed changes.

Here is an example video. https://streamable.com/unr9s
